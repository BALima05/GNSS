import os
import json
import re
import datetime
import zipfile
import io
from pathlib import Path

# Arquivo onde salvaremos o "estado" atual (qual pasta está sendo processada)
ARQUIVO_ESTADO = "sessao_atual.json"

def salvar_estado(chave, valor):
    """Salva uma informação para ser usada pelo próximo script."""
    dados = {}
    if os.path.exists(ARQUIVO_ESTADO):
        try:
            with open(ARQUIVO_ESTADO, 'r') as f:
                dados = json.load(f)
        except: pass
    
    dados[chave] = str(valor) # Converte Path para string
    
    with open(ARQUIVO_ESTADO, 'w') as f:
        json.dump(dados, f, indent=4)
    print(f"💾 Estado salvo: {chave} -> {valor}")

def carregar_estado(chave):
    """Recupera uma informação salva anteriormente."""
    if not os.path.exists(ARQUIVO_ESTADO):
        return None
    try:
        with open(ARQUIVO_ESTADO, 'r') as f:
            dados = json.load(f)
            return dados.get(chave)
    except:
        return None

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