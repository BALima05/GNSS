import os
import zipfile
import shutil
import subprocess
import re
import concurrent.futures
from pathlib import Path

# MAX_ZIP_SIZE foi removida, pois usaremos o RTKLIB diretamente
# Adicionar uma forma de capturar os caminhos na máquina posteriormente!! (pendente)
crx_path = r"C:\Users\berna\Downloads\RNXCMP_4.2.0_Windows_mingw_64bit\RNXCMP_4.2.0_Windows_mingw_64bit\CRX2RNX.exe"
teqc_path = r"C:\Users\berna\OneDrive\Documentos\GNSS\Programas\teqc.exe"

# Função para imprimir a etapa atual do processamento
def print_etapa(etapa):
    print(f"\n{'='*40}\n[ETAPA] {etapa}\n{'='*40}")

def descompactar_zip(origem_path, pasta_destino_d_path):
    """
    Descompacta arquivos ZIP de origem.
    Procura por zips aninhados, extrai tudo e move apenas os arquivos .d
    para a pasta de destino final.
    """
    origem_path = Path(origem_path)
    pasta_destino_d_path = Path(pasta_destino_d_path)
    
    # Cria diretórios temporários na pasta base
    temp_raiz = pasta_destino_d_path.parent / "TEMP_ZIPS"
    temp_extraidos = pasta_destino_d_path.parent / "TEMP_EXTRAIDOS"
    
    os.makedirs(temp_raiz, exist_ok=True)
    os.makedirs(temp_extraidos, exist_ok=True)
    os.makedirs(pasta_destino_d_path, exist_ok=True)

    try:
        if origem_path.is_file() and origem_path.suffix.lower() == ".zip":
            print(f"🗃️ Extraindo pacote principal ZIP: {origem_path.name}")
            with zipfile.ZipFile(origem_path, 'r') as zip_ref:
                zip_ref.extractall(temp_raiz)
        elif origem_path.is_dir():
            print(f"📂 Copiando arquivos ZIP da pasta: {origem_path}")
            for arquivo in os.listdir(origem_path):
                if arquivo.lower().endswith('.zip'):
                    shutil.copy(origem_path / arquivo, temp_raiz / arquivo)
        else:
            print(f"❌ Erro: Caminho de origem não é um arquivo .zip ou diretório válido.")
            return

        # Extrai os zips individuais (que podem conter os arquivos .d)
        for arquivo_zip in temp_raiz.glob('*.zip'):
            print(f" extracting... {arquivo_zip.name}")
            try:
                with zipfile.ZipFile(arquivo_zip, 'r') as zip_ref:
                    zip_ref.extractall(temp_extraidos)
                print(f"✅ Descompactado: {arquivo_zip.name}")
            except zipfile.BadZipFile:
                print(f"❌ ZIP inválido: {arquivo_zip.name}")
        
        # Procura recursivamente por arquivos .d e move
        regex_d = re.compile(r".*\.\d{2}d$", re.IGNORECASE)
        for raiz, _, arquivos in os.walk(temp_extraidos):
            for arquivo in arquivos:
                if arquivo.lower().endswith('.d') or regex_d.match(arquivo):
                    origem = Path(raiz) / arquivo
                    destino = pasta_destino_d_path / arquivo
                    shutil.move(origem, destino)
                    print(f"📁 Movido: {arquivo} -> {pasta_destino_d_path.name}")
    
    finally:
        # Limpa os diretórios temporários
        shutil.rmtree(temp_raiz, ignore_errors=True)
        shutil.rmtree(temp_extraidos, ignore_errors=True)
        print("🧹 Limpeza temporária concluída.")

def _processar_crx(arquivo_d_path, crx2rnx_path):
    """Função auxiliar para paralelismo do CRX2RNX."""
    comando = f'"{crx2rnx_path}" "{arquivo_d_path}"'
    try:
        subprocess.run(comando, shell=True, check=True, cwd=arquivo_d_path.parent,
                         capture_output=True, text=True)
        nome_saida = arquivo_d_path.with_suffix("." + arquivo_d_path.suffix[1:3] + "o").name
        return f"🔁 Convertido: {arquivo_d_path.name} → {nome_saida}"
    except subprocess.CalledProcessError as e:
        return f"❌ Erro ao converter: {arquivo_d_path.name}. (Verifique CRX2RNX e permissões)\n{e.stderr}"

def converter_crx2rnx(pasta_d_path, crx2rnx_path):
    """Converte arquivos .d para .o em paralelo."""
    regex_d = re.compile(r".*\.\d{2}d$", re.IGNORECASE)
    arquivos_d = [f for f in pasta_d_path.glob('*') if f.is_file() and regex_d.match(f.name)]
    
    if not arquivos_d:
        print("❌ Nenhum arquivo .d válido (ex: .22d) encontrado para conversão.")
        return

    print(f"Iniciando conversão de {len(arquivos_d)} arquivos Hatanaka...")
    
    # Usa ProcessPoolExecutor para rodar várias instâncias do CRX2RNX ao mesmo tempo
    with concurrent.futures.ProcessPoolExecutor() as executor:
        # Cria uma "tarefa" para cada arquivo, passando o caminho do CRX2RNX
        tarefas = {executor.submit(_processar_crx, arquivo_d, crx2rnx_path): arquivo_d for arquivo_d in arquivos_d}
        
        # Coleta os resultados à medida que ficam prontos
        for futuro in concurrent.futures.as_completed(tarefas):
            print(futuro.result())

def _processar_teqc(arquivo_o_path, teqc_path, gps_dir, glonass_dir, gps_glonass_dir):
    """Função auxiliar para paralelismo do TEQC."""
    try:
        arquivo = arquivo_o_path.name
        gps_saida = gps_dir / f"GPS_{arquivo}"
        glonass_saida = glonass_dir / f"GLONASS_{arquivo}"
        gps_glonass_saida = gps_glonass_dir / f"GPS_GLONASS_{arquivo}"

        # -R = GPS, -E = Galileo/BeiDou/QZSS (excluir)
        subprocess.run(f'"{teqc_path}" -R -E "{arquivo_o_path}" > "{gps_saida}"', shell=True, check=True)
        # -G = GLONASS
        subprocess.run(f'"{teqc_path}" -G -E "{arquivo_o_path}" > "{glonass_saida}"', shell=True, check=True)
        # Padrão (GPS+GLONASS)
        subprocess.run(f'"{teqc_path}" -E "{arquivo_o_path}" > "{gps_glonass_saida}"', shell=True, check=True)
        
        return f"🛰️  Processado TEQC: {arquivo}"
    except subprocess.CalledProcessError as e:
        return f"❌ Erro no TEQC: {arquivo_o_path.name}. O arquivo pode estar corrompido.\n{e}"

def separar_teqc(pasta_d_path, pasta_saida_path, teqc_path):
    """Separa arquivos .o por constelação em paralelo."""
    gps_dir = pasta_saida_path / "GPS"
    glonass_dir = pasta_saida_path / "GLONASS"
    gps_glonass_dir = pasta_saida_path / "GPS_GLONASS"
    os.makedirs(gps_dir, exist_ok=True)
    os.makedirs(glonass_dir, exist_ok=True)
    os.makedirs(gps_glonass_dir, exist_ok=True)

    regex_o = re.compile(r".*\.\d{2}o$", re.IGNORECASE)
    arquivos_o = [f for f in pasta_d_path.glob('*') if regex_o.match(f.name)]
    
    if not arquivos_o:
        print("❌ Nenhum arquivo .o encontrado para processamento de satélite!")
        return
    
    print(f"Iniciando separação por satélite de {len(arquivos_o)} arquivos...")

    with concurrent.futures.ProcessPoolExecutor() as executor:
        tarefas = {executor.submit(_processar_teqc, arquivo_o, teqc_path, gps_dir, glonass_dir, gps_glonass_dir): arquivo_o for arquivo_o in arquivos_o}
        
        for futuro in concurrent.futures.as_completed(tarefas):
            print(futuro.result())

# --- FUNÇÃO 'compactar_por_lote' REMOVIDA ---

def main():
    print("🔧 PROCESSAMENTO GNSS - SCRIPT OTIMIZADO (PARA RTKLIB)")
    
    origem_zip = input("📂 Caminho do .ZIP do IBGE ou pasta com zips: ").strip().strip('"')
    pasta_base = input("📁 Caminho onde deseja salvar os dados processados: ").strip().strip('"')
    mes_ano = input("🗓️ Informe o mês e ano (ex: NOV_22): ").strip().strip('"').upper()

    # Validação dos executáveis
    CAMINHO_CRX2RNX = Path(crx_path)
    CAMINHO_TEQC = Path(teqc_path)
    
    if not CAMINHO_CRX2RNX.is_file():
        print(f"❌ Erro: CRX2RNX.exe não encontrado em '{CAMINHO_CRX2RNX}'")
        return
    if not CAMINHO_TEQC.is_file():
        print(f"❌ Erro: teqc.exe não encontrado em '{CAMINHO_TEQC}'")
        return

    # Usa Pathlib para gerenciar pastas
    pasta_final = Path(pasta_base) / mes_ano
    os.makedirs(pasta_final, exist_ok=True)

    pasta_d   = pasta_final / "1 - Dados tipos .d"
    pasta_sep = pasta_final / "2 - Dados separados por satélite (Prontos para RTKLIB)"
    # --- pasta_zip FOI REMOVIDA ---

    print_etapa("1/3 - Descompactando e separando arquivos .d")
    descompactar_zip(origem_zip, pasta_d)

    print_etapa("2/3 - Convertendo Hatanaka (.d) p/ RINEX (.o) [EM PARALELO]")
    converter_crx2rnx(pasta_d, CAMINHO_CRX2RNX)

    print_etapa("3/3 - Separando arquivos por satélite (TEQC) [EM PARALELO]")
    separar_teqc(pasta_d, pasta_sep, CAMINHO_TEQC)

    print_etapa("🎉 FINALIZAÇÃO")
    print(f"Processamento concluído! Seus arquivos RINEX estão prontos para o RTKLIB em:")
    print(f"{pasta_sep}")

if __name__ == "__main__":
    main()