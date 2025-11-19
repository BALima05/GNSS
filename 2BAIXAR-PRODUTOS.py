import os
import ftplib
import datetime
import gnsscal
import gzip
import shutil
from pathlib import Path
from tqdm import tqdm

# Configurações do Servidor FTP do GFZ (Geralmente aberto/anônimo)
FTP_HOST = "ftp.gfz-potsdam.de"
FTP_BASE_PATH = "/GNSS/products"  # Caminho base dos produtos

def descompactar_z_gz(caminho_arquivo):
    """Descompacta arquivos .Z ou .gz para que o RTKLIB possa ler."""
    caminho_arquivo = Path(caminho_arquivo)
    # Remove a extensão de compressão para o nome final
    caminho_final = caminho_arquivo.with_suffix('')
    
    print(f"   🔓 Descompactando: {caminho_arquivo.name}...")
    
    try:
        # Tenta usar gzip (funciona para .gz e muitas vezes para .Z modernos)
        with gzip.open(caminho_arquivo, 'rb') as f_in:
            with open(caminho_final, 'wb') as f_out:
                shutil.copyfileobj(f_in, f_out)
        
        # Remove o arquivo compactado original para economizar espaço
        os.remove(caminho_arquivo)
        return caminho_final
    except Exception as e:
        print(f"   ⚠️ Falha ao descompactar automaticamente (pode ser necessário 7zip): {e}")
        return caminho_arquivo

def baixar_arquivo_ftp(ftp, pasta_remota, nome_arquivo, pasta_local):
    """Baixa um arquivo específico do FTP com barra de progresso."""
    caminho_local = Path(pasta_local) / nome_arquivo
    
    try:
        tamanho_arquivo = ftp.size(f"{pasta_remota}/{nome_arquivo}")
    except:
        tamanho_arquivo = 0

    print(f"   ⬇️ Baixando: {nome_arquivo}")
    
    with open(caminho_local, 'wb') as f:
        with tqdm(total=tamanho_arquivo, unit='B', unit_scale=True, desc=nome_arquivo, leave=False) as pbar:
            def callback(data):
                f.write(data)
                pbar.update(len(data))
            
            try:
                ftp.retrbinary(f"RETR {pasta_remota}/{nome_arquivo}", callback)
                return caminho_local
            except ftplib.error_perm as e:
                print(f"   ❌ Erro: Arquivo não encontrado no servidor: {e}")
                f.close()
                os.remove(caminho_local)
                return None

def buscar_e_baixar_produtos(data_alvo, pasta_saida):
    """
    Lógica principal:
    1. Converte Data -> Semana GPS
    2. Conecta no FTP
    3. Tenta achar arquivos SP3 e CLK (nomes curtos ou longos)
    """
    pasta_saida = Path(pasta_saida)
    os.makedirs(pasta_saida, exist_ok=True)

    # 1. Cálculos de Tempo
    semana_gps, dia_semana = gnsscal.date2gpswd(data_alvo)
    print(f"\n🌍 Processando Data: {data_alvo} | Semana GPS: {semana_gps} | Dia: {dia_semana}")

    # 2. Conexão FTP
    try:
        ftp = ftplib.FTP(FTP_HOST)
        ftp.login() # Login anônimo
        print(f"   ✅ Conectado a {FTP_HOST}")
    except Exception as e:
        print(f"   ❌ Falha na conexão FTP: {e}")
        return

    # Caminho da semana: /GNSS/products/{semana}
    pasta_remota = f"{FTP_BASE_PATH}/{semana_gps}"
    
    try:
        ftp.cwd(pasta_remota)
    except:
        print(f"   ❌ Pasta da semana {semana_gps} não encontrada no servidor.")
        ftp.quit()
        return

    # 3. Definir nomes de arquivos para procurar
    # O RTKLIB gosta de nomes curtos: igsWWWD.sp3
    # O Servidor pode ter nomes longos: IGS0OPSFIN...
    
    # Tentativa 1: Nomes Curtos (Padrão Antigo - Mais compatível com scripts simples)
    arquivos_alvo = [
        f"igs{semana_gps}{dia_semana}.sp3.Z",       # Órbita
        f"igs{semana_gps}{dia_semana}.clk_30s.Z",   # Relógio 30s (Melhor)
        f"igs{semana_gps}{dia_semana}.clk.Z"        # Relógio 5min (Fallback)
    ]

    # Listar arquivos na pasta para ver o que tem
    arquivos_no_servidor = []
    try:
        arquivos_no_servidor = ftp.nlst()
    except:
        pass

    for alvo in arquivos_alvo:
        # Verifica se o arquivo curto existe direto
        if alvo in arquivos_no_servidor:
            arquivo_baixado = baixar_arquivo_ftp(ftp, pasta_remota, alvo, pasta_saida)
            if arquivo_baixado:
                descompactar_z_gz(arquivo_baixado)
        else:
            # Se não achou o curto, tenta achar o Longo equivalente
            # Lógica simplificada: Procura algo que tenha o dia do ano ou semana
            # (Isso é complexo de fazer perfeito, então focamos no .Z padrão que o GFZ mantém)
            print(f"   ⚠️ Arquivo {alvo} não encontrado explicitamente.")

    ftp.quit()
    print("   ✅ Download da data finalizado.")

def main():
    print("🛰️ DOWNLOADER DE PRODUTOS IGS (GFZ FTP)")
    
    pasta_destino = input("📂 Pasta para salvar os produtos (ex: C:\\GNSS\\PRODUTOS): ").strip().strip('"')
    
    # Modo de entrada: Data única ou Intervalo? Vamos fazer simples por enquanto.
    data_str = input("🗓️ Data do levantamento (DD/MM/AAAA): ").strip()
    
    try:
        dia, mes, ano = map(int, data_str.split('/'))
        data_alvo = datetime.date(ano, mes, dia)
        
        buscar_e_baixar_produtos(data_alvo, pasta_destino)
        
        print(f"\n🎉 Arquivos prontos em: {pasta_destino}")
        print("DICA: Aponte esta pasta no script de processamento PPP anterior.")
        
    except ValueError:
        print("❌ Formato de data inválido.")

if __name__ == "__main__":
    main()