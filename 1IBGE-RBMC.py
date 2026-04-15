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

# Descobrir Mês e Ano automaticamente pelo ZIP
def descobrir_mes_ano_automatico(caminho_origem):
    """
    Analisa ZIP/Pasta (e Zips aninhados) para encontrar o primeiro arquivo GNSS
    válido e determinar o Mês e Ano.
    """
    caminho = Path(caminho_origem)
    print(f"\n🕵️ MODO RAIO-X: Analisando data em {caminho.name}...")

    # Regex: (Dia)(Sessão).(Ano)(Tipo) -> ex: 0010.24d, 001a.24o
    padrao_data = re.compile(r"(\d{3})[0-9a-zA-Z]\.(\d{2})[dDoO]")

    def tentar_extrair_data(lista_nomes):
        """Helper para verificar uma lista de nomes de arquivos"""
        for nome in lista_nomes:
            nome_limpo = Path(nome).name
            match = padrao_data.search(nome_limpo)
            if match:
                doy = int(match.group(1))
                ano = int(match.group(2))
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
            for f in origem_path.glob("*.zip"):
                shutil.copy(f, temp_raiz / f.name)
        
        print(">> Extraindo zips internos...")
        zips_internos = list(temp_raiz.glob('*.zip'))
        for arq_zip in zips_internos:
            try:
                with zipfile.ZipFile(arq_zip, 'r') as z:
                    z.extractall(temp_extraidos)
            except zipfile.BadZipFile:
                print(f"⚠️ Aviso: Zip corrompido: {arq_zip.name}")

        print(">> Organizando arquivos (.d e navegação)...")
        count_d = 0
        count_n = 0
        
        for raiz, dirs, arquivos in os.walk(temp_extraidos):
            for arquivo in arquivos:
                caminho_origem = Path(raiz) / arquivo
                
                # ####### LÓGICA DE SEPARAÇÃO #######
                
                # 1. Arquivos Hatanaka (.YYd)
                if re.search(r"\.\d{2}d$", arquivo, re.IGNORECASE):
                    shutil.move(caminho_origem, pasta_destino_d_path / arquivo)
                    count_d += 1
                    
                # 2. Arquivos de Navegação (.YYn = GPS, .YYg = GLONASS, .YYp = Misto)
                elif re.search(r"\.\d{2}[ngp]$", arquivo, re.IGNORECASE):
                    shutil.move(caminho_origem, pasta_destino_nav_path / arquivo)
                    count_n += 1

        print(f"✅ Extração concluída: {count_d} arquivos .d e {count_n} arquivos de navegação.")

    finally:
        # Limpeza
        if temp_raiz.exists(): shutil.rmtree(temp_raiz)
        if temp_extraidos.exists(): shutil.rmtree(temp_extraidos)

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
    
    origem_zip = config.IBGE_ZIP
    mes_ano = descobrir_mes_ano_automatico(origem_zip)

    # Validação dos executáveis
    CAMINHO_CRX2RNX = Path(config.CRX2RNX_PATH)
    CAMINHO_TEQC = Path(config.TEQC_PATH)
    
    if not CAMINHO_CRX2RNX.is_file():
        print(f"❌ Erro: CRX2RNX.exe não encontrado em '{CAMINHO_CRX2RNX}'")
        return
    if not CAMINHO_TEQC.is_file():
        print(f"❌ Erro: teqc.exe não encontrado em '{CAMINHO_TEQC}'")
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

    print_etapa("2/3 - Convertendo Hatanaka (.d) p/ RINEX (.o) [EM PARALELO]")
    converter_crx2rnx(pasta_d, CAMINHO_CRX2RNX)

    print_etapa("3/3 - Separando arquivos por satélite (TEQC) [EM PARALELO]")
    separar_teqc(pasta_d, pasta_sep, CAMINHO_TEQC)

    print_etapa("🎉 FINALIZAÇÃO")
    print(f"Processamento concluído! Seus arquivos RINEX estão prontos para o RTKLIB em:")
    print(f"{pasta_sep}")

    # Salva o caminho das pastas para o proximo script
    utils.salvar_estado("pasta_rinex_pronta", pasta_sep)
    utils.salvar_estado("mes_ano", mes_ano)

if __name__ == "__main__":
    main()