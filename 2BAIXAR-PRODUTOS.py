import os
import ftplib
import gzip
import shutil
import datetime
import subprocess
from pathlib import Path

# CONFIGURAÇÕES
SERVER_IGS = "igs.ign.fr"  # Mirror do IGS
PATH_IGS = "/pub/igs/products" # Caminho base no FTP

# Caminho para o 7-Zip
PATH_7ZIP = r"C:\Program Files\7-Zip\7z.exe" 

def gps_date_converter(year, doy):
    """
    Converte Ano e Dia do Ano (DOY) para Semana GPS e Dia da Semana GPS.
    Retorna: (gps_week, gps_day_of_week, data_datetime)
    """
    # Ajuste do ano (assumindo 20xx)
    full_year = 2000 + int(year) if int(year) < 100 else int(year)
    
    # Data do arquivo
    date_obj = datetime.datetime(full_year, 1, 1) + datetime.timedelta(days=int(doy) - 1)
    
    # Data epoch do GPS (06/01/1980)
    gps_epoch = datetime.datetime(1980, 1, 6)
    
    # Cálculo
    delta = date_obj - gps_epoch
    gps_week = delta.days // 7
    gps_dow = delta.days % 7
    
    return gps_week, gps_dow, date_obj

def download_igs_product(ftp, gps_week, gps_dow, pasta_destino):
    """
    Baixa arquivos .sp3.Z e .clk.Z do servidor FTP.
    """
    # Nome dos arquivos (Padrão IGS Final: igsWWWD.sp3.Z)
    # Ex: igs22342.sp3.Z (Semana 2234, Dia 2)
    base_name = f"igs{gps_week}{gps_dow}"
    files_to_download = [f"{base_name}.sp3.Z", f"{base_name}.clk.Z"]
    
    # Tenta entrar na pasta da semana
    try:
        ftp.cwd(f"{PATH_IGS}/{gps_week}")
    except ftplib.error_perm:
        print(f"❌ Erro: Pasta da semana {gps_week} não encontrada no servidor.")
        return []

    downloaded_files = []

    for filename in files_to_download:
        local_path = pasta_destino / filename
        
        # Se o arquivo já existe descompactado ou compactado, pula
        final_file = pasta_destino / filename[:-2] # Remove .Z
        if final_file.exists():
            print(f"🔹 Arquivo já existe (ignorado): {final_file.name}")
            continue
            
        print(f"⬇️ Baixando: {filename} ...")
        try:
            with open(local_path, "wb") as f:
                ftp.retrbinary(f"RETR {filename}", f.write)
            downloaded_files.append(local_path)
            print(f"✅ Download concluído: {filename}")
        except ftplib.error_perm:
            print(f"⚠️ Arquivo não encontrado no servidor: {filename}")
            if local_path.exists(): os.remove(local_path)
    
    # Volta para a raiz para a próxima iteração
    ftp.cwd("/")
    return downloaded_files

def descompactar_z(arquivo_path):
    """
    Descompacta arquivos .Z usando 7-Zip ou tenta gzip (se for renomeado).
    No Windows, .Z é chato de abrir nativamente com Python puro.
    """
    arquivo_path = Path(arquivo_path)
    if not arquivo_path.exists(): return

    # Verifica se é .Z
    if arquivo_path.suffix == '.Z':
        # Tenta usar 7-Zip se configurado
        if os.path.exists(PATH_7ZIP):
            cmd = [PATH_7ZIP, 'e', str(arquivo_path), f'-o{str(arquivo_path.parent)}', '-y']
            subprocess.run(cmd, capture_output=True)
            # Remove o .Z após extrair
            os.remove(arquivo_path)
            print(f"📦 Descompactado (7-Zip): {arquivo_path.name}")
        else:
            print(f"⚠️ AVISO: Não foi possível descompactar {arquivo_path.name} automaticamente.")
            print(f"   Instale o 7-Zip e ajuste o caminho no script, ou descompacte manualmente.")
            print(f"   O RTKLIB precisa dos arquivos .sp3 e .clk (sem o .Z).")

def main():
    print("🌍 DOWNLOAD AUTOMÁTICO DE EFEMÉRIDES IGS (ORBITS/CLOCKS)")
    print("-------------------------------------------------------")
    
    # 1. Obter pasta dos RINEX (.o)
    pasta_rinex = input("📂 Pasta onde estão os arquivos RINEX (.o): ").strip().strip('"')
    pasta_rinex = Path(pasta_rinex)
    
    if not pasta_rinex.exists():
        print("❌ Pasta não encontrada.")
        return

    # 2. Preparar pasta de destino dos produtos
    pasta_produtos = pasta_rinex.parent / "IGS_PRODUCTS"
    os.makedirs(pasta_produtos, exist_ok=True)
    print(f"📂 Os produtos serão salvos em: {pasta_produtos}")

    # 3. Escanear datas necessárias
    # Padrão RINEX 2: ssssdddh.yyo (Ex: .22o)
    arquivos_o = list(pasta_rinex.glob("*.*o"))
    
    if not arquivos_o:
        print("❌ Nenhum arquivo RINEX encontrado.")
        return

    datas_necessarias = set()

    print("\n🔎 Analisando arquivos para determinar datas...")
    for arq in arquivos_o:
        # Tenta extrair do nome do arquivo (assumindo padrão ssssDDDH.YYo)
        try:
            # Pega a extensão (ex: .22o) -> Ano = 22
            ano = int(arq.suffix[1:3])
            # Pega os caracteres do nome para o dia (ex: nome do arquivo 'brft0220.22o' -> dia 022)
            # Se o arquivo vier do IBGE/TEQC, geralmente os ultimos 4 chars antes do ponto são DDDH
            nome_puro = arq.stem # brft0220
            doy = int(nome_puro[-4:-1]) # Pega 022
            
            week, dow, date_obj = gps_date_converter(ano, doy)
            datas_necessarias.add((week, dow))
            print(f"   📄 {arq.name} -> Data: {date_obj.strftime('%d/%m/%Y')} (Semana GPS: {week}, Dia: {dow})")
        except Exception as e:
            print(f"   ⚠️ Não foi possível ler data de {arq.name}. Verifique se segue o padrão 'nomeDDDH.yyo'.")

    if not datas_necessarias:
        print("❌ Nenhuma data válida identificada.")
        return

    # 4. Conectar e Baixar
    print(f"\n📡 Conectando ao servidor FTP {SERVER_IGS}...")
    try:
        ftp = ftplib.FTP(SERVER_IGS)
        ftp.login() # Login anônimo
        print("✅ Conexão estabelecida!")
        
        for week, dow in datas_necessarias:
            z_files = download_igs_product(ftp, week, dow, pasta_produtos)
            
            # 5. Descompactar
            for z_file in z_files:
                descompactar_z(z_file)
                
        ftp.quit()
        print("\n🎉 Downloads finalizados!")
        print(f"Certifique-se de usar a pasta abaixo no script '3RTKlib-PPP.py':")
        print(f"{pasta_produtos}")

    except Exception as e:
        print(f"\n❌ Erro de conexão ou download: {e}")

if __name__ == "__main__":
    main()