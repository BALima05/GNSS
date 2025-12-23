import os
import ftplib
import datetime
import re
import gzip
import shutil
from pathlib import Path

# =============================================================================
# CONFIGURAÇÕES
# =============================================================================
SERVER_IGS = "igs.ign.fr"
PATH_IGS = "/pub/igs/products"

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

def decompress_gz(file_path):
    """Descompacta arquivos .gz nativamente."""
    new_path = file_path.with_suffix('') # Remove .gz
    with gzip.open(file_path, 'rb') as f_in:
        with open(new_path, 'wb') as f_out:
            shutil.copyfileobj(f_in, f_out)
    os.remove(file_path) # Remove o original compactado
    print(f"📦 Descompactado: {new_path.name}")
    return new_path

def download_igs_smart(ftp, gps_week, gps_dow, date_obj, pasta_destino):
    """
    Tenta baixar arquivos usando o padrão NOVO (Longo) e cai para o ANTIGO (Curto) se falhar.
    """
    # Dados para montagem dos nomes
    yyyy = date_obj.strftime("%Y")
    ddd = date_obj.strftime("%j") # Dia do ano 001-366
    
    # --- DEFINIÇÃO DOS NOMES DE ARQUIVO ---
    # 1. Padrão NOVO (IGS0OPSFIN...) - Usado pós-2022
    # Orbita: IGS0OPSFIN_YYYYDDD0000_01D_15M_ORB.SP3.gz
    orb_long = f"IGS0OPSFIN_{yyyy}{ddd}0000_01D_15M_ORB.SP3.gz"
    # Relógio: Tenta 30s (melhor) primeiro, depois 05m
    clk_long_30s = f"IGS0OPSFIN_{yyyy}{ddd}0000_01D_30S_CLK.CLK.gz"
    clk_long_05m = f"IGS0OPSFIN_{yyyy}{ddd}0000_01D_05M_CLK.CLK.gz"

    # 2. Padrão ANTIGO (igsWWWD...) - Usado pré-2022
    base_short = f"igs{gps_week}{gps_dow}"
    orb_short = f"{base_short}.sp3.Z"
    clk_short = f"{base_short}.clk.Z"

    # Muda para a pasta da semana GPS
    try:
        ftp.cwd(f"{PATH_IGS}/{gps_week}")
    except ftplib.error_perm:
        print(f"❌ Erro: Pasta da semana {gps_week} não encontrada.")
        return

    # Função auxiliar de download
    def try_download(filenames_priority_list):
        for fname in filenames_priority_list:
            local_path = pasta_destino / fname
            final_path = pasta_destino / fname.replace('.gz', '').replace('.Z', '')
            
            # Se já existe descompactado, pula
            if final_path.exists():
                print(f"🔹 Já existe: {final_path.name}")
                return True
            
            print(f"⬇️ Tentando: {fname} ...")
            try:
                with open(local_path, "wb") as f:
                    ftp.retrbinary(f"RETR {fname}", f.write)
                print(f"✅ Sucesso: {fname}")
                
                # Descompactar
                if fname.endswith('.gz'):
                    decompress_gz(local_path)
                elif fname.endswith('.Z'):
                    # Para .Z antigo, avisa usuário (ou use 7zip se tiver configurado, 
                    # mas o foco aqui é 2024 que usa .gz)
                    print(f"⚠️ Arquivo .Z baixado. Use 7-Zip para extrair: {fname}")
                return True
            except ftplib.error_perm:
                if local_path.exists(): os.remove(local_path)
                continue # Tenta o próximo da lista
        return False

    # --- BAIXAR ORBITAS ---
    if not try_download([orb_long, orb_short]):
        print(f"❌ Falha ao baixar Órbitas para {date_obj.date()}")

    # --- BAIXAR RELÓGIOS ---
    if not try_download([clk_long_30s, clk_long_05m, clk_short]):
        print(f"❌ Falha ao baixar Relógios para {date_obj.date()}")

    ftp.cwd("/") # Volta para raiz

def main():
    print("🌍 DOWNLOAD IGS V3 (Suporte a nomes Longos/2024+)")
    
    pasta_rinex = input("📂 Pasta onde estão os arquivos RINEX (.o): ").strip().strip('"')
    pasta_rinex = Path(pasta_rinex)
    
    if not pasta_rinex.exists():
        print("❌ Pasta não encontrada.")
        return

    pasta_produtos = pasta_rinex.parent / "IGS_PRODUCTS"
    os.makedirs(pasta_produtos, exist_ok=True)
    
    arquivos_o = list(pasta_rinex.glob("*.*o"))
    datas_processar = set()

    print("\n🔎 Identificando datas...")
    for arq in arquivos_o:
        ano, doy = extrair_info_arquivo(arq)
        if ano is not None:
            wk, dw, dt = gps_date_converter(ano, doy)
            datas_processar.add((wk, dw, dt))
            print(f"   📅 {dt.date()} (Semana {wk}) <- {arq.name}")
    
    if not datas_processar:
        print("❌ Nenhuma data válida encontrada.")
        return

    print("\n📡 Conectando ao FTP IGN...")
    try:
        ftp = ftplib.FTP(SERVER_IGS)
        ftp.login()
        
        for wk, dw, dt in sorted(datas_processar):
            print(f"\n--- Processando {dt.date()} (Semana {wk}) ---")
            download_igs_smart(ftp, wk, dw, dt, pasta_produtos)
            
        ftp.quit()
        print("\n🎉 Todos os downloads concluídos!")
        print(f"📂 Arquivos salvos em: {pasta_produtos}")

    except Exception as e:
        print(f"❌ Erro crítico: {e}")

if __name__ == "__main__":
    main()