import os
import ftplib
import datetime
import re
import gzip
import shutil
from pathlib import Path
import utils
import config
import subprocess
import fnmatch

# =============================================================================
# CONFIGURAÇÕES
# =============================================================================
SERVER_IGS = "igs.ign.fr"
PATH_IGS = "/pub/igs/products"

# Tenta localizar o 7-Zip automaticamente (necessário para arquivos .Z)
PATH_7ZIP = None
caminhos_comuns = [
    r"C:\Program Files\7-Zip\7z.exe",
    r"C:\Program Files (x86)\7-Zip\7z.exe"
]
for p in caminhos_comuns:
    if os.path.exists(p):
        PATH_7ZIP = p
        break

def gps_date_converter(year, doy):
    """Converte Ano e DOY para Semana GPS, Dia da Semana e Data completa."""
    full_year = 2000 + int(year) if int(year) < 100 else int(year)
    date_obj = datetime.datetime(full_year, 1, 1) + datetime.timedelta(days=int(doy) - 1)
    gps_epoch = datetime.datetime(1980, 1, 6)
    delta = date_obj - gps_epoch
    gps_week = delta.days // 7
    gps_dow = delta.days % 7
    return gps_week, gps_dow, date_obj

def extrair_info_arquivo(arquivo_path):
    """Extrai Ano e DOY do nome do arquivo RINEX."""
    try:
        extensao = arquivo_path.suffix
        ano = int(extensao[1:3])
        nome = arquivo_path.stem
        # Regex para pegar os 3 dígitos do dia antes do último caractere
        match = re.search(r'(\d{3}).$', nome)
        if match:
            doy = int(match.group(1))
            return ano, doy
        return None, None
    except:
        return None, None
    
def descompactar_z_7zip(arquivo_z, pasta_destino):
    """Usa o 7-Zip para descompactar arquivos .Z antigos."""
    if not PATH_7ZIP:
        return False
    try:
        # Comando: 7z e "arquivo" -o"destino" -y (sobrescrever)
        cmd = [PATH_7ZIP, "e", str(arquivo_z), f"-o{pasta_destino}", "-y"]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except:
        return False
    
def encontrar_melhor_arquivo(lista_arquivos_server, ano, doy, tipo):
    """
    Procura na lista do servidor o melhor arquivo para a data,
    dando PRIORIDADE MÁXIMA aos arquivos originais do IGS (IGS0OPS),
    exigindo OBRIGATORIAMENTE que o arquivo de relógio seja de 30 Segundos (30S)
    para que o PPP convirja corretamente.
    Ignora propositalmente produtos de outros centros (como o CODE) 
    que causam incompatibilidade de leitura no RTKLIB.
    """
    ano_str = str(ano)
    doy_str = f"{int(doy):03d}" # Ex: 004
    
    candidatos = []

    # Padrão 1: Nome Longo (Prioridade IGS Moderno)
    # Ex: IGS0OPSFIN_20240040000_01D_15M_ORB.SP3.gz
    # Busca por *YYYYDDD*ORB.SP3*
    if tipo == 'sp3':
        padrao_longo = f"*_{ano_str}{doy_str}*ORB.SP3*" 
    else: # clk
        # Forçar o "30S" no nome do arquivo
        padrao_longo = f"*_{ano_str}{doy_str}*30S_CLK.CLK*"

    # Padrão 2: Nome Curto (Legado)
    # Ex: igsWWWD.sp3.Z (Isso é difícil de montar aqui sem a semana, 
    # então confia-se na busca por substring se o longo falhar)
    
    for f in lista_arquivos_server:
        if fnmatch.fnmatch(f, padrao_longo):
            candidatos.append(f)
    
    # Se achou nomes longos, tenta pegar o "FIN" (Final) em vez do "RAP" (Rapid)
    if candidatos:
        finais = [c for c in candidatos if "FIN" in c]
        if finais: return finais[0] # Retorna o primeiro Final encontrado
        return candidatos[0] # Se não tiver Final, vai o Rapid mesmo
    
    # Se retornar None, a função main() já está programada para 
    # acionar o "Fallback" e baixar os arquivos de nome curto (igsWWWD.sp3)
    return None

def main():
    print("🌍 DOWNLOAD IGS V3 (Suporte a nomes Longos/2024+)")

    if PATH_7ZIP:
        print(f"🔧 7-Zip detectado em: {PATH_7ZIP}")
    else:
        print("⚠️ 7-Zip não encontrado. Se houver apenas arquivos .Z, eles não serão abertos.")
    
    caminho_salvo = utils.carregar_estado("pasta_rinex_pronta")
    pasta_rinex = None

    if caminho_salvo:
        pasta_rinex = Path(caminho_salvo)
        print(f"📂 Base recuperada: {pasta_rinex}")
    else:
        #Fallback caso não tenha sido salvo
        print("⚠️ Memória não encontrada. Recalculando...")
        mes_ano = utils.descobrir_mes_ano_automatico(config.IBGE_ZIP)
        pasta_rinex = Path(config.PASTA_BASE) / mes_ano / "2 - Dados separados por satélite (Prontos para RTKLIB)"

    # -- Correção de segurança --
    # Se por acaso o script 1 salvou o caminho de uma subpasta (ex.: .../GPS),
    # subimos 1 nível para garantir que estamos na pasta mãe
    if pasta_rinex.name in ["GPS", "GLONASS", "GPS_GLONASS"]:
        pasta_rinex = pasta_rinex.parent
        print(f"⚠️ Ajuste de caminho: Subpasta detectada. Usando pasta mãe: {pasta_rinex}")
    
    if not pasta_rinex.exists():
        print(f"❌ A pasta {pasta_rinex} não foi encontrada.")
        return

    pasta_produtos = pasta_rinex.parent / "IGS_PRODUCTS"
    os.makedirs(pasta_produtos, exist_ok=True)
    utils.salvar_estado("pasta_produtos", pasta_produtos)
    
    print("\n🔍 Varrendo todas as subpastas (GPS, GLONASS, GPS_GLONASS)")

    # rglob procura recursivamente em todas as pastas
    arquivos_o = list(pasta_rinex.rglob("*.*o"))
    datas_processar = set()

    print("\n🔎 Identificando datas...")
    for arq in arquivos_o:
        ano, doy = extrair_info_arquivo(arq)
        if ano is not None:
            wk, dw, dt = gps_date_converter(ano, doy)
            full_year = dt.year
            datas_processar.add((wk, dw, dt, full_year, doy))
            print(f"   📅 {dt.date()} (Semana {wk}) <- {arq.name}")
    
    if not datas_processar:
        print("❌ Nenhuma data válida encontrada.")
        return

    print(f"✅ Encontrados {len(datas_processar)} arquivos RINEX distribuídos nas pastas.")
    print(f"✅ Necessário baixar produtos para {len(datas_processar)} dias distintos.")

    print("\n📡 Conectando ao FTP IGN...")
    try:
        ftp = ftplib.FTP(SERVER_IGS)
        ftp.login()
        # Ativa modo passivo para evitar bloqueios de firewall
        ftp.set_pasv(True)
        print("✅ Conexão estabelecida.")

        lista_ordenada = sorted(list(datas_processar), key=lambda x: x[2]) # Ordena por semana
        
        for wk, dw, dt, full_year, doy in lista_ordenada:
            print(f"   ⬇️ Processando dia: {dt.date()} (Week {wk})")
            pasta_remota = f"{PATH_IGS}/{wk}"
            
            # Lista de alvos (Orbitas e Relógios)
            alvos = [
                f"igs{wk}{dw}.sp3.gz",     # Preferido
                f"igs{wk}{dw}.sp3.Z",      # Fallback (precisa de 7-Zip)
                f"igs{wk}{dw}.clk_30s.gz", # Prioridade relógio preciso
                f"igs{wk}{dw}.clk_30s.Z",
                f"igs{wk}{dw}.clk.gz",
                f"igs{wk}{dw}.clk.Z"
            ]

            try:
                ftp.cwd(pasta_remota)
                arquivos_no_server = ftp.nlst()
                
                # Procura Arquivos (SP3 e CLK)
                alvos_encontrados = []

                # Agrupa alvos por tipo (SP3 ou CLK) para não baixar repetido
                sp3_baixado = False
                clk_baixado = False

                # 1. Busca SP3 (Efemerides)
                arquivo_sp3 = encontrar_melhor_arquivo(arquivos_no_server, full_year, doy, 'sp3')
                if arquivo_sp3:
                    alvos_encontrados.append(arquivo_sp3)
                else:
                    # Tenta fallback nome curto
                    curto = f"igs{wk}{dw}.sp3.Z"
                    if curto in arquivos_no_server: alvos_encontrados.append(curto)
                    else: print(f"   ⚠️ SP3 não encontrado para {dt.date()}")

                # 2. Busca CLK (Relógios)
                arquivo_clk = encontrar_melhor_arquivo(arquivos_no_server, full_year, doy, 'clk')
                if arquivo_clk:
                    alvos_encontrados.append(arquivo_clk)
                else:
                    # Tenta fallback nome curto
                    curto_clk = f"igs{wk}{dw}.clk.Z" # ou clk_30s.Z
                    if curto_clk in arquivos_no_server: alvos_encontrados.append(curto_clk)

                # --- DOWNLOAD ---
                for arquivo_remoto in alvos_encontrados:
                    local_file = pasta_produtos / arquivo_remoto
                    
                    # Nome final esperado após descompactar (remove .gz ou .Z)
                    nome_limpo = arquivo_remoto.replace('.gz', '').replace('.Z', '')
                    final_path = pasta_produtos / nome_limpo

                    if final_path.exists() and final_path.stat().st_size > 0:
                        print(f"   ⏩ Já existe: {nome_limpo}")
                        continue

                    print(f"   ⬇️  Baixando: {arquivo_remoto} ... ", end="")
                    try:
                        with open(local_file, 'wb') as f:
                            ftp.retrbinary(f'RETR {arquivo_remoto}', f.write)
                        print("OK!")
                    except Exception as e:
                        print(f"ERRO: {e}")
                        if local_file.exists(): os.remove(local_file)

            except ftplib.error_perm:
                print(f"❌ Pasta remota {pasta_remota} não encontrada.")

        ftp.quit()
        print("\n✅ Download concluído.")

    except Exception as e:
        print(f"❌ Erro FTP: {e}")
        return

    # -- Descompactação -- 
    print("\n📦 Verificando arquivos compactados...")

    # 1. GZIP (nativo no Python)
    for arq_gz in pasta_produtos.glob("*.gz"):
        try:
            with gzip.open(arq_gz, 'rb') as f_in:
                with open(arq_gz.with_suffix(''), 'wb') as f_out:
                    shutil.copyfileobj(f_in, f_out)
            os.remove(arq_gz)
            print(f"   🔨 GZIP Extraído: {arq_gz.name}")
        except Exception as e:
            print(f"   ❌ Erro no GZIP {arq_gz.name}: {e}")
    
    # Tratamento para .Z (se houver ferramenta externa ou renomear)
    # 2. .Z (Unix Compress) - Usa 7-Zip se disponível
    z_files = list(pasta_produtos.glob("*.Z"))
    if z_files:
        if PATH_7ZIP:
            print(f"   ⚙️ Usando 7-Zip para {len(z_files)} arquivos .Z...")
            for arq_z in z_files:
                if descompactar_z_7zip(arq_z, pasta_produtos):
                    print(f"      🔓 Extraído: {arq_z.name}")
                    os.remove(arq_z) # Remove o original se deu certo
                else:
                    print(f"      ❌ Falha ao extrair: {arq_z.name}")
        else:
            print(f"   ⚠️ ATENÇÃO: {len(z_files)} arquivos .Z não foram extraídos (Instale o 7-Zip).")
            
    print(f"\n📂 Produtos prontos em: {pasta_produtos}")

if __name__ == "__main__":
    main()