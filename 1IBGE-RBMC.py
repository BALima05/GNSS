import os
import zipfile
import shutil
import subprocess
import re
import concurrent.futures
from pathlib import Path
import config
import datetime
import io
import utils
import urllib.request
import urllib.error
import gzip
import time

# Descobrir Mês e Ano automaticamente pelo ZIP
def descobrir_mes_ano_automatico(caminho_origem):
    """
    Analisa ZIP/Pasta (e Zips aninhados) para encontrar o primeiro arquivo GNSS
    válido (RINEX 2 ou RINEX 3) e determinar o Mês e Ano.
    """
    caminho = Path(caminho_origem)
    print(f"\n🕵️ MODO RAIO-X: Analisando data em {caminho.name}...")

    # Padrão RINEX 2: (Dia)(Sessão).(Ano)(Tipo) -> ex: 0010.24d, poli001a.24o
    padrao_rnx2 = re.compile(r"(\d{3})[0-9a-zA-Z]\.(\d{2})[dDoO]")

    # Padrão RINEX 3: _(Ano)(DOY)(HoraMinuto)_ -> ex: POLI00BRA_R_20240310000_01D...
    padrao_rnx3 = re.compile(r"_(\d{4})(\d{3})\d{4}_")

    def tentar_extrair_data(lista_nomes):
        """Helper para verificar uma lista de nomes de arquivos"""
        for nome in lista_nomes:
            nome_limpo = Path(nome).name
            
            match3 = padrao_rnx3.search(nome_limpo)
            if match3:
                ano_completo = int(match3.group(1)) # Pega os 4 dígitos do Ano
                doy = int(match3.group(2))          # Pega os 3 dígitos do DOY
                data_obj = datetime.datetime(ano_completo, 1, 1) + datetime.timedelta(days=doy - 1)
                return data_obj.strftime("%b_%y").upper(), nome_limpo

            # 2. Tenta identificar como formato clássico (RINEX 2)
            match2 = padrao_rnx2.search(nome_limpo)
            if match2:
                doy = int(match2.group(1))          # Pega os 3 dígitos do DOY
                ano = int(match2.group(2))          # Pega os 2 dígitos do Ano
                ano_completo = 2000 + ano
                data_obj = datetime.datetime(ano_completo, 1, 1) + datetime.timedelta(days=doy - 1)
                return data_obj.strftime("%b_%y").upper(), nome_limpo
                
        return None, None

    try:
        # CASO 1: A origem é um ARQUIVO .ZIP
        if caminho.is_file() and caminho.suffix.lower() == '.zip':
            with zipfile.ZipFile(caminho, 'r') as z_main:
                # 1. Tenta achar na raiz do ZIP principal
                res, arq = tentar_extrair_data(z_main.namelist())
                if res:
                    print(f"✅ Data encontrada no Nível 1: {res} (Arquivo: {arq})")
                    return res
                
                # 2. Se não achou, procura dentro dos ZIPS internos (Nível 2)
                print("   ↳ Nível 1 sem arquivos de dados. Olhando dentro dos Zips internos...")
                for item in z_main.namelist():
                    if item.lower().endswith('.zip'):
                        try:
                            # Abre o zip interno na memória (sem extrair pro disco)
                            with z_main.open(item) as zip_bytes:
                                with zipfile.ZipFile(zip_bytes) as z_nested:
                                    res_n, arq_n = tentar_extrair_data(z_nested.namelist())
                                    if res_n:
                                        print(f"✅ Data encontrada no Nível 2 ({item}): {res_n} (Arquivo: {arq_n})")
                                        return res_n
                        except:
                            continue # Se um zip interno der erro, pula pro próximo

        # CASO 2: A origem é uma PASTA
        elif caminho.is_dir():
            # 1. Tenta achar arquivos soltos na pasta
            arquivos_pasta = [f.name for f in caminho.glob('*')]
            res, arq = tentar_extrair_data(arquivos_pasta)
            if res: return res

            # 2. Tenta olhar dentro dos Zips que estão na pasta
            zips_na_pasta = list(caminho.glob('*.zip'))
            print(f"   ↳ Verificando {len(zips_na_pasta)} zips dentro da pasta...")
            for zip_path in zips_na_pasta:
                try:
                    with zipfile.ZipFile(zip_path, 'r') as z:
                        res, arq = tentar_extrair_data(z.namelist())
                        if res:
                            print(f"✅ Data encontrada dentro de {zip_path.name}: {res}")
                            return res
                except: continue

    except Exception as e:
        print(f"❌ Erro ao ler estrutura de arquivos: {e}")

    print("\n⚠️ AVISO: Não foi possível determinar a data automaticamente.")
    print("   -> Usando data atual como fallback.")
    return datetime.datetime.now().strftime("%b_%y").upper()

# MAX_ZIP_SIZE foi removida, pois usaremos o RTKLIB diretamente

# Função para imprimir a etapa atual do processamento
def print_etapa(etapa):
    print(f"\n{'='*40}\n[ETAPA] {etapa}\n{'='*40}")

def descompactar_zip(origem_path, pasta_destino_d_path, pasta_destino_nav_path):
    """
    Descompacta arquivos ZIP e separa:
    - .d -> pasta_destino_d
    - .n, .g, .p -> pasta_destino_nav
    """
    origem_path = Path(origem_path) # Pasta de origem
    pasta_destino_d_path = Path(pasta_destino_d_path) # Pasta destino para arquivos .d
    pasta_destino_nav_path = Path(pasta_destino_nav_path) # Pasta destino para arquivos .nav
    
    # Cria diretórios temporários na pasta base
    temp_raiz = pasta_destino_d_path.parent / "TEMP_ZIPS"
    temp_extraidos = pasta_destino_d_path.parent / "TEMP_EXTRAIDOS"
    
    # Cria os diretórios temporários
    os.makedirs(temp_raiz, exist_ok=True)
    os.makedirs(temp_extraidos, exist_ok=True)
    os.makedirs(pasta_destino_d_path, exist_ok=True)
    os.makedirs(pasta_destino_nav_path, exist_ok=True)

    try:
        print(">> Copiando arquivos ZIP para pasta temporária...")
        if origem_path.is_file() and origem_path.suffix.lower() == ".zip":
            with zipfile.ZipFile(origem_path, 'r') as z:
                z.extractall(temp_raiz)
        elif origem_path.is_dir():
            zips_origem = [f for f in origem_path.iterdir() if f.is_file() and f.suffix.lower() == '.zip']
            for f in zips_origem:
                shutil.copy(f, temp_raiz / f.name)
        
        print(">> Extraindo zips internos...")
        zips_internos = [f for f in temp_raiz.iterdir() if f.is_file() and f.suffix.lower() == '.zip']
        for arq_zip in zips_internos:
            try:
                with zipfile.ZipFile(arq_zip, 'r') as z:
                    z.extractall(temp_extraidos)
            except zipfile.BadZipFile:
                print(f"⚠️ Aviso: Zip corrompido: {arq_zip.name}")

        print(">> Organizando arquivos (.d e navegação)...")
        count_d = 0
        count_n = 0

        pastas_para_buscar = [temp_raiz, temp_extraidos]
        
        for pasta in pastas_para_buscar:
            if not pasta.exists():
                continue
        
            for raiz, dirs, arquivos in os.walk(pasta):
                for arquivo in arquivos:
                    caminho_origem = Path(raiz) / arquivo
                    arq_lower = arquivo.lower()
                
                    # ####### LÓGICA DE SEPARAÇÃO #######
                
                    # 1. RINEX 2 (.24d) ou RINEX 3 (.crx ou .crx.gz)
                    if re.search(r"\.\d{2}d$", arq_lower) or arq_lower.endswith(".crx") or arq_lower.endswith(".crx.gz"):
                        if caminho_origem.exists(): # Evita tentar mover algo que já foi movido
                            shutil.move(caminho_origem, pasta_destino_d_path / arquivo)
                            count_d += 1
                        
                    # 2. Arquivos de Navegação (.YYn, .YYg, .YYp ou RINEX 3 _MN.rnx / _MN.rnx.gz)
                    elif re.search(r"\.\d{2}[ngp]$", arq_lower) or ("_mn" in arq_lower and "rnx" in arq_lower):
                        if caminho_origem.exists():
                            shutil.move(caminho_origem, pasta_destino_nav_path / arquivo)
                            count_n += 1

        print(f"✅ Extração concluída: {count_d} arquivos de observação e {count_n} arquivos de navegação.")

    finally:
        # Limpeza
        if temp_raiz.exists(): shutil.rmtree(temp_raiz, ignore_errors=True)
        if temp_extraidos.exists(): shutil.rmtree(temp_extraidos, ignore_errors=True)

def _processar_crx(arquivo_d_path, crx2rnx_path):
    """Converte Hatanaka em RINEX usando RNXCMP (Suporta Rinex 2 e 3)"""
    try:
        # Se for um .crx.gz (RINEX 3 Compactado), extrai primeiro
        if arquivo_d_path.suffix.lower() == '.gz':
            import gzip
            with gzip.open(arquivo_d_path, 'rb') as f_in:
                novo_path = arquivo_d_path.with_suffix('')
                with open(novo_path, 'wb') as f_out:
                    shutil.copyfileobj(f_in, f_out)
            os.remove(arquivo_d_path)
            arquivo_d_path = novo_path

        # O CRX2RNX é inteligente o suficiente para saber se a saída será .o ou .rnx
        cmd = f'"{crx2rnx_path}" "{arquivo_d_path}"'
        resultado = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        
        if resultado.returncode == 0:
            os.remove(arquivo_d_path)
            return f"✅ Convertido: {arquivo_d_path.name}"
        else:
            return f"❌ Erro na conversão de {arquivo_d_path.name}: {resultado.stderr}"
    except Exception as e:
        return f"❌ Erro fatal em {arquivo_d_path.name}: {e}"

def converter_crx2rnx_paralelo(pasta_d_path, crx2rnx_path):
    # Pega tanto os .d antigos quanto os .crx modernos
    arquivos_d = [
        f for f in pasta_d_path.iterdir() 
        if f.is_file() and (re.search(r"\.\d{2}[dD]$", f.name) or ".crx" in f.name.lower())
    ]
    
    with concurrent.futures.ProcessPoolExecutor() as executor:
        tarefas = {executor.submit(_processar_crx, arq, crx2rnx_path): arq for arq in arquivos_d}
        for futuro in concurrent.futures.as_completed(tarefas):
            print(futuro.result())

def _processar_gfzrnx(arquivo_o_path, gfzrnx_path, gps_dir, glonass_dir, gps_glonass_dir):
    """Função auxiliar para paralelismo do GFZRNX (O substituto do TEQC)."""
    try:
        arquivo = arquivo_o_path.name
        gps_saida = gps_dir / f"GPS_{arquivo}"
        glonass_saida = glonass_dir / f"GLONASS_{arquivo}"
        gps_glonass_saida = gps_glonass_dir / f"GPS_GLONASS_{arquivo}"
        
        # GFZRNX: -satsys = satélite manipular. 'G' = GPS, 'R' = GLONASS
        # Cria arquivo só de GPS
        subprocess.run(f'"{gfzrnx_path}" -finp "{arquivo_o_path}" -fout "{gps_saida}" -satsys G', shell=True, check=True)
        
        # Cria arquivo só de GLONASS
        subprocess.run(f'"{gfzrnx_path}" -finp "{arquivo_o_path}" -fout "{glonass_saida}" -satsys R', shell=True, check=True)
        
        # Cria arquivo GPS + GLONASS
        subprocess.run(f'"{gfzrnx_path}" -finp "{arquivo_o_path}" -fout "{gps_glonass_saida}" -satsys GR', shell=True, check=True)
        
        return f"✅ GFZRNX fatiou: {arquivo}"
    except Exception as e:
        return f"❌ Erro GFZRNX em {arquivo}: {e}"

def separar_constelacoes_paralelo(pasta_d_path, pasta_sep_path, gfzrnx_path):
    gps_dir = pasta_sep_path / "GPS"
    glonass_dir = pasta_sep_path / "GLONASS"
    gps_glonass_dir = pasta_sep_path / "GPS_GLONASS"
    
    os.makedirs(gps_dir, exist_ok=True)
    os.makedirs(glonass_dir, exist_ok=True)
    os.makedirs(gps_glonass_dir, exist_ok=True)

    # Coleta arquivos .o (Rinex 2) e .rnx (Rinex 3)
    arquivos_obs = [
        f for f in pasta_d_path.iterdir() 
        if f.is_file() and (re.search(r"\.\d{2}[oO]$", f.name) or f.name.lower().endswith(".rnx"))
    ]
    
    with concurrent.futures.ProcessPoolExecutor() as executor:
        tarefas = {executor.submit(_processar_gfzrnx, arq, gfzrnx_path, gps_dir, glonass_dir, gps_glonass_dir): arq for arq in arquivos_obs}
        for futuro in concurrent.futures.as_completed(tarefas):
            print(futuro.result())

# --- FUNÇÃO 'compactar_por_lote' REMOVIDA ---

def baixar_navegacao_brdc(ano, doy, pasta_destino_nav, max_tentativas=5):
    """
    Baixa o arquivo de navegação global Multi-GNSS (BRDC) do servidor IGS/BKG.
    Exemplo de URL: https://igs.bkg.bund.de/root_ftp/IGS/BRDC/2024/001/BRDC00IGS_R_20240010000_01D_MN.rnx.gz
    """
    pasta_destino_nav = Path(pasta_destino_nav)
    os.makedirs(pasta_destino_nav, exist_ok=True)
    
    ano_str = str(ano)
    doy_str = str(doy).zfill(3) # Garante 3 dígitos (ex: 001, 045)
    
    nome_arquivo_gz = f"BRDC00IGS_R_{ano_str}{doy_str}0000_01D_MN.rnx.gz"
    nome_arquivo_rnx = f"BRDC00IGS_R_{ano_str}{doy_str}0000_01D_MN.rnx"
    
    caminho_gz = pasta_destino_nav / nome_arquivo_gz
    caminho_rnx = pasta_destino_nav / nome_arquivo_rnx
    
    # Se já existir, não baixa de novo
    if caminho_rnx.exists():
        print(f"✅ Navegação BRDC já existe para o dia {doy}: {nome_arquivo_rnx}")
        return caminho_rnx
        
    url = f"https://igs.bkg.bund.de/root_ftp/IGS/BRDC/{ano_str}/{doy_str}/{nome_arquivo_gz}"

    for tentativa in range(1, max_tentativas + 1):
        print(f"🌐 Baixando navegação global BRDC do dia {doy}/{ano}...")
    
        try:
            # Baixa com timeout de 15 segundos para evitar que o programa trave
            with urllib.request.urlopen(url, timeout=15) as response, open(caminho_gz, 'wb') as out_file:
                shutil.copyfileobj(response, out_file)
            
            # Extrai o arquivo
            with gzip.open(caminho_gz, 'rb') as f_in:
                with open(caminho_rnx, 'wb') as f_out:
                    shutil.copyfileobj(f_in, f_out)
                    
            os.remove(caminho_gz) # Apaga o .gz
            print("✅ Navegação baixada e extraída com sucesso!")
            return caminho_rnx
            
        except urllib.error.URLError as e:
            print(f"   ⚠️ Falha na rede (Tentativa {tentativa}): {e}")
            if caminho_gz.exists(): os.remove(caminho_gz) # Limpa download corrompido
            
            if tentativa < max_tentativas:
                print("   ⏳ Aguardando 5 segundos antes de tentar novamente...")
                time.sleep(5)
            else:
                print(f"   ❌ Erro definitivo ao baixar navegação do dia {doy}/{ano}.")
                return None
        except Exception as e:
            print(f"   ❌ Erro inesperado no dia {doy}/{ano}: {e}")
            if caminho_gz.exists(): os.remove(caminho_gz)
            return None
    
def resolver_navegacao(pasta_d, pasta_nav):
    """ Vasculha os arquivos descompactados e baixa as navegações necessárias. """
    padrao_rnx2 = re.compile(r"(\d{3})[0-9a-zA-Z]\.(\d{2})[dDoO]")
    padrao_rnx3 = re.compile(r"_(\d{4})(\d{3})\d{4}_")

    dias_processados = set()

    # Inspeciona os arquivos na pasta para descobrir os dias (Ano, DOY)
    for arquivo in pasta_d.iterdir():
        if not arquivo.is_file(): continue
        
        nome = arquivo.name
        ano, doy = None, None
        
        match3 = padrao_rnx3.search(nome)
        if match3:
            ano = int(match3.group(1))
            doy = int(match3.group(2))
        else:
            match2 = padrao_rnx2.search(nome)
            if match2:
                doy = int(match2.group(1))
                ano = 2000 + int(match2.group(2))
                
        if ano and doy:
            dias_processados.add((ano, doy))

    if not dias_processados:
        print("⚠️ Nenhum dia válido encontrado para baixar navegação.")
        return

    # Baixa a navegação para cada dia único encontrado
    for ano, doy in sorted(list(dias_processados)):
        baixar_navegacao_brdc(ano, doy, pasta_nav)

def main():
    print("🔧 PROCESSAMENTO GNSS - SCRIPT OTIMIZADO (PARA RTKLIB)")
    
    origem_zip = config.IBGE_ZIP
    mes_ano = descobrir_mes_ano_automatico(origem_zip)

    # Validação dos executáveis
    CAMINHO_CRX2RNX = Path(config.CRX2RNX_PATH)
    CAMINHO_GFZRNX = Path(config.GFZRNX_PATH)
    
    if not CAMINHO_CRX2RNX.is_file():
        print(f"❌ Erro: CRX2RNX não encontrado em '{CAMINHO_CRX2RNX}'")
        return
    if not CAMINHO_GFZRNX.is_file():
        print(f"❌ Erro: GFZRNX não encontrado em '{CAMINHO_GFZRNX}'")
        return

    # usa Pathlib para gerenciar pastas
    pasta_final = Path(config.PASTA_BASE) / mes_ano
    os.makedirs(pasta_final, exist_ok=True)

    pasta_d   = pasta_final / "1 - Dados tipos .d"
    pasta_nav = pasta_final / "1.1 - Navegacao Broadcast"
    pasta_sep = pasta_final / "2 - Dados separados por satélite (Prontos para RTKLIB)"
    # pasta_zip FOI REMOVIDA

    print_etapa("1/3 - Descompactando e separando arquivos .d")
    descompactar_zip(origem_zip, pasta_d, pasta_nav)

    print_etapa("1.5/3 - Obtendo órbitas globais (Navegação BRDC)")
    resolver_navegacao(pasta_d, pasta_nav)

    print_etapa("2/3 - Convertendo Hatanaka (.d) p/ RINEX (.o) [EM PARALELO]")
    converter_crx2rnx_paralelo(pasta_d, CAMINHO_CRX2RNX)

    print_etapa("3/3 - Separando arquivos por satélite (GFZRNX) [EM PARALELO]")
    separar_constelacoes_paralelo(pasta_d, pasta_sep, CAMINHO_GFZRNX)

    print_etapa("🎉 FINALIZAÇÃO")
    print(f"Processamento concluído! Seus arquivos RINEX estão prontos para o RTKLIB em:")
    print(f"{pasta_sep}")

    # Salva o caminho das pastas para o proximo script
    utils.salvar_estado("pasta_rinex_pronta", pasta_sep)
    utils.salvar_estado("mes_ano", mes_ano)

if __name__ == "__main__":
    main()