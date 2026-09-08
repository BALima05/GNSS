import os
import subprocess
from pathlib import Path
import concurrent.futures
import config
import utils
import re
import datetime

def detectar_sistemas(arquivo_obs):
    """
    Lê o cabeçalho do RINEX de observação e retorna o conjunto de sistemas
    presentes (letras: G=GPS, R=GLONASS, E=Galileo, C=BeiDou, J=QZSS, I=IRNSS, S=SBAS).
    Funciona tanto para RINEX 3.x ("SYS / # / OBS TYPES") quanto RINEX 2.x
    (usa o indicador de sistema da linha "RINEX VERSION / TYPE").
    """
    sistemas = set()
    try:
        with open(arquivo_obs, 'r', errors='ignore') as f:
            for _ in range(300):  # cabeçalho não deve passar disso
                linha = f.readline()
                if not linha:
                    break
                if "END OF HEADER" in linha:
                    break
                if "RINEX VERSION / TYPE" in linha:
                    # coluna ~41 (index 40) traz o indicador de sistema: G/R/E/J/C/I/S/M(mixed)
                    indicador = linha[40:41].strip().upper()
                    if indicador and indicador != 'M':
                        sistemas.add(indicador)
                elif "SYS / # / OBS TYPES" in linha:
                    # RINEX 3: primeiro caractere da linha é o código do sistema
                    codigo = linha[0:1].strip().upper()
                    if codigo:
                        sistemas.add(codigo)
    except Exception as e:
        print(f"⚠️ Não consegui ler cabeçalho de {arquivo_obs.name} para detectar sistemas: {e}")

    return sistemas


def navsys_bitmask(sistemas):
    """Converte o conjunto de sistemas detectados no valor pos1-navsys do RTKLIB."""
    mapa = {'G': 1, 'S': 2, 'R': 4, 'E': 8, 'J': 16, 'C': 32, 'I': 64}
    valor = 0
    for s in sistemas:
        valor |= mapa.get(s, 0)
    # Fallback de segurança: se não detectou nada, mantém GPS+GLONASS (comportamento antigo)
    return valor if valor > 0 else 5


def gerar_config_com_navsys(config_file, navsys_valor, pasta_saida, nome_base):
    """
    Cria uma cópia temporária do .conf substituindo (ou adicionando) a linha
    pos1-navsys pelo valor detectado automaticamente para este arquivo.
    """
    config_file = Path(config_file)
    conf_temp = Path(pasta_saida) / f"_tmp_{nome_base}.conf"

    linha_encontrada = False
    with open(config_file, 'r', encoding='utf-8', errors='ignore') as f_in:
        linhas = f_in.readlines()

    with open(conf_temp, 'w', encoding='utf-8') as f_out:
        for linha in linhas:
            if linha.strip().startswith('pos1-navsys'):
                f_out.write(f"pos1-navsys={navsys_valor}\n")
                linha_encontrada = True
            else:
                f_out.write(linha)
        if not linha_encontrada:
            f_out.write(f"pos1-navsys={navsys_valor}\n")

    return conf_temp


def extrair_ano_doy(nome_arq):
    """
    Extrai (ano, doy) do nome de um arquivo RINEX/produto IGS, tentando
    o padrão RINEX 3 (_YYYYDDDhhmm_, usado em obs, nav, sp3 e clk) e,
    se falhar, o padrão RINEX 2 (nnnDDDa.YYo).
    Retorna (None, None) se não encontrar.
    """
    match3 = re.search(r"_(\d{4})(\d{3})\d{4}_", nome_arq)
    if match3:
        return int(match3.group(1)), int(match3.group(2))

    match2 = re.search(r"^[a-zA-Z0-9]{4}(\d{3})[a-zA-Z0-9]\.(\d{2})[oO]$", nome_arq)
    if match2:
        return 2000 + int(match2.group(2)), int(match2.group(1))

    return None, None


def gps_date_converter(year, doy):
    """Converte o ano e DOY para a semana GPS (usada nos produtos IGS)."""
    full_year = 2000 + int(year) if int(year) < 100 else int(year)
    date_obj = datetime.datetime(full_year, 1, 1) + datetime.timedelta(days=int(doy) - 1)
    gps_epoch = datetime.datetime(1980, 1, 6)
    delta = date_obj - gps_epoch
    gps_week = delta.days // 7
    gps_dow = delta.days % 7
    return gps_week, gps_dow, full_year, int(doy)

def processar_ppp_rtklib(arquivo_obs, pasta_produtos, pasta_nav, config_file, rnx2rtkp_path, pasta_saida):
    try:
        arquivo_obs = Path(arquivo_obs)
        pasta_produtos = Path(pasta_produtos)
        pasta_nav = Path(pasta_nav)
        pasta_saida = Path(pasta_saida)
        
        # Nome do arquivo de saída
        arquivo_pos = pasta_saida / arquivo_obs.with_suffix('.pos').name
        
        # Obtenção dos arquivos nav e produtos
        nav_files_brutos = list(pasta_nav.glob("*.[0-9][0-9]n")) + \
                           list(pasta_nav.glob("*.[0-9][0-9]p")) + \
                           list(pasta_nav.glob("*.[0-9][0-9]g")) + \
                           list(pasta_nav.glob("*.nav")) + \
                           list(pasta_nav.glob("*MN.rnx"))
                           
        nav_files = []
        for nav in nav_files_brutos:
            # RTKLIB ignora navegação terminada em .rnx. Vamos forçar para .nav
            if nav.suffix.lower() == '.rnx':
                novo_nav = nav.with_suffix('.nav')
                nav.rename(novo_nav) # Renomeia fisicamente no HD
                nav_files.append(novo_nav)
            else:
                nav_files.append(nav)
        
        arquivos_sp3 = list(pasta_produtos.glob("*.[sS][pP]3")) + list(pasta_produtos.glob("*.eph"))
        arquivos_clk = list(pasta_produtos.glob("*.[cC][lL][kK]"))
        # Arquivos de bias/DCB (corrige o viés P1-C1 entre código C/A e código P,
        # relevante quando o receptor não rastreia P1 mas os produtos precisos
        # (SP3/CLK) foram gerados com base em P1/P2). Aceita tanto arquivos
        # datados (ex.: COD0MGXFIN_20240050000_01D_01D_OSB.BIA) quanto arquivos
        # estáticos genéricos do IGS/CODE (ex.: P1C1_ALL.DCB, P1P2_ALL.DCB), que
        # não têm data no nome e por isso são sempre incluídos quando presentes.
        arquivos_bias = list(pasta_produtos.glob("*.[bB][iI][aA]")) + list(pasta_produtos.glob("*.[dD][cC][bB]"))

        # --- LÓGICA DE EXTRAÇÃO DE DATA (comparação EXATA, não substring) ---
        # IMPORTANTE: usar "doy_str in nome_arquivo" é perigoso, porque o DOY
        # pode coincidir por acaso com parte do ANO no nome do arquivo
        # (ex.: dia 024 é substring de "2024"), fazendo o filtro pegar
        # arquivos de navegação de dias errados. Por isso,
        # extraímos (ano, doy) de cada arquivo com regex e comparamos
        # igualdade exata com o (ano, doy) do arquivo de observação.
        ano, doy_int = extrair_ano_doy(arquivo_obs.name)

        if ano and doy_int:
            wk, dw, full_year, _ = gps_date_converter(ano % 100, doy_int)

            # Filtra os de Navegação comparando (ano, doy) exatamente
            nav_files = [
                f for f in nav_files
                if extrair_ano_doy(f.name) == (ano, doy_int)
            ]

            # Filtra SP3 e CLK: aceita tanto o nome-longo (ano+doy exatos)
            # quanto o padrão curto de nome de produto IGS por semana GPS (igsWWWD)
            padrao_curto = f"igs{wk}{dw}"
            arquivos_sp3 = [
                f for f in arquivos_sp3
                if extrair_ano_doy(f.name) == (ano, doy_int) or padrao_curto in f.name
            ]
            arquivos_clk = [
                f for f in arquivos_clk
                if extrair_ano_doy(f.name) == (ano, doy_int) or padrao_curto in f.name
            ]

            # Bias/DCB: se o arquivo tem data no nome, exige data exata (igual a nav/sp3/clk).
            # Se não tem data (arquivo genérico do tipo P1C1_ALL.DCB), mantém sempre.
            arquivos_bias = [
                f for f in arquivos_bias
                if extrair_ano_doy(f.name) == (ano, doy_int)
                or extrair_ano_doy(f.name) == (None, None)
            ]
        else:
            return f"⚠️ Pulei {arquivo_obs.name}: Não consegui identificar a data no nome do arquivo."
        # ---------------------------------------

        if not arquivos_sp3:
            return f"⚠️ Pulei {arquivo_obs.name}: Faltam arquivos .sp3 (Orbitas)."
        if not arquivos_clk:
            return f"⚠️ Pulei {arquivo_obs.name}: Faltam arquivos .clk (Relógios)."
        if not nav_files:
            return f"⚠️ Pulei {arquivo_obs.name}: Faltam arquivos de navegação (.n, .p, .g, .nav)."
            
        # --- DETECÇÃO AUTOMÁTICA DE NAVSYS ---
        # lê o cabeçalho do arquivo de observação pra saber quais
        # sistemas (GPS/GLONASS/...) existem nele, gerando um .conf
        # temporário com o pos1-navsys correto. Isso evita usar
        # navsys=GPS+GLONASS num arquivo que só tem GLONASS ou vice-versa
        sistemas_detectados = detectar_sistemas(arquivo_obs)
        navsys_valor = navsys_bitmask(sistemas_detectados)
        config_efetivo = gerar_config_com_navsys(
            config_file, navsys_valor, pasta_saida, arquivo_obs.stem
        )
        print(f"🛰️  {arquivo_obs.name}: sistemas={sorted(sistemas_detectados) or '??'} -> navsys={navsys_valor}")

        # Monta o comando
        cmd = [
            str(rnx2rtkp_path),
            '-k', str(config_efetivo),
            '-x', '3',  # Nível 3 de Trace (Gera arquivo .trace detalhado)
            '-y', '3',  # Nível 3 de Status (Gera arquivo .stat)
            '-o', str(arquivo_pos),
            str(arquivo_obs)
        ]
        
        # Adiciona produtos e navegação ao comando

        for nav in nav_files:      # OBRIGATÓRIO PASSAR NAV PRIMEIRO
            cmd.append(str(nav))
        for sp3 in arquivos_sp3:
            cmd.append(str(sp3))
        for clk in arquivos_clk:
            cmd.append(str(clk))
        for bias in arquivos_bias:  # Opcional: corrige DCB P1-C1 quando disponível
            cmd.append(str(bias))

        if arquivos_bias:
            print(f"   ↳ Bias/DCB aplicado: {[b.name for b in arquivos_bias]}")
        else:
            print(f"   ⚠️ Nenhum arquivo de bias/DCB encontrado para {arquivo_obs.name} "
                  f"(ok se o receptor rastreia P1/P2 nativamente, senão pode inserir viés de código)")

        # Executa capturando TUDO
        try:
            result = subprocess.run(cmd, capture_output=True, text=True)

            if result.returncode != 0:
                return (
                    f"Falha no RTKLIB para {arquivo_obs.name}\n"
                    f"stderr: {result.stderr}\n"
                    f"stdout: {result.stdout}"
                )

            # --- Verificar se o arquivo EXISTE e tem CONTEÚDO ---
            if arquivo_pos.exists() and arquivo_pos.stat().st_size > 0:
                return f"✅ PPP Sucesso: {arquivo_pos.name} (Tamanho: {arquivo_pos.stat().st_size/1024:.1f} KB)"
            else:
                # Se o arquivo não foi criado, mostra o erro que o RTKLIB cuspiu
                return (f"❌ Falha {arquivo_obs.name} (Arquivo vazio ou não criado).\n"
                        f"   Log RTKLIB: {result.stderr}\n"
                        f"   Output: {result.stdout}")
        finally:
            # Limpeza do arquivo de configuração temporário
            try:
                config_efetivo.unlink(missing_ok=True)
            except Exception as e:
                print(f"⚠️ Não foi possível remover conf temporário {config_efetivo}: {e}")

    except Exception as e:
        return f"💥 Erro de execução processando {arquivo_obs.name}: {e}"

def main():
    print("🌍 AUTOMÇÃO DE PPP COM RTKLIB (Python Wrapper)")
    
    # --- CONFIGURAÇÕES ---
    # Caminho para o executável rnx2rtkp.exe
    path_rnx2rtkp = Path(config.RNX2RTKP_PATH)
    if not path_rnx2rtkp.is_file():
        print(f"❌ Executável não encontrado: {path_rnx2rtkp}")
        return
    
    # Pasta onde estão os arquivos RINEX .o (gerados no script anterior)
    path_rinex_obs = utils.carregar_estado("pasta_rinex_pronta")
    
    # Pasta onde salvou os arquivos .sp3 e .clk baixados do IGS
    path_produtos = utils.carregar_estado("pasta_produtos")

    # Pasta onde estão os dados de navegação
    path_nav = None
    
    # Arquivo de configuração .conf
    path_config = Path(config.CONFIG_FILE)
    if not path_config.is_file():
        print(f"❌ Arquivo de configuração não encontrado: {path_config}")
        return
    
    # Pasta para salvar os resultados
    path_saida = os.path.join(os.path.dirname(path_rinex_obs), "RESULTADOS_PPP")
    os.makedirs(path_saida, exist_ok=True)

    # --- VERIFICAÇÕES ---

    if path_rinex_obs:
        path_rinex_obs = Path(path_rinex_obs)
        print(f"📁 Base recuperada: {path_rinex_obs}")
    else:
        print("⚠️ Memória não encontrada. Recalculando...")
        mes_ano = utils.descobrir_mes_ano_automatico(config.IBGE_ZIP)
        path_rinex_obs = Path(config.PASTA_BASE) / mes_ano / "2 - Dados separados por satélite (Prontos para RTKLIB)"
    
    # -- Correção de segurança --
    # Se por acaso o script 1 salvou o caminho de uma subpasta (ex.: .../GPS),
    # subimos 1 nível para garantir que estamos na pasta mãe
    if path_rinex_obs.name in ["GPS", "GLONASS", "GPS_GLONASS"]:
        path_rinex_obs = path_rinex_obs.parent
        print(f"⚠️ Ajuste de caminho: Subpasta detectada. Usando pasta mãe: {path_rinex_obs}")
    
    if not path_rinex_obs.exists():
        print(f"❌ A pasta {path_rinex_obs} não foi encontrada.")
        return
    
    # Recuperando a pasta nav diretamente
    path_nav = path_rinex_obs.parent / "1.1 - Navegacao Broadcast"
    if not path_nav.exists():
        print(f"⚠️ Pasta de navegação não encontrada em {path_nav}.")
        return

    # --- PROCESSAMENTO ---
    print("\n🔍 Varrendo todas as subpastas (GPS, GLONASS, GPS_GLONASS)")
    arquivos_o = [
        arquivo
        for arquivo in path_rinex_obs.rglob("*")
        if arquivo.is_file() and arquivo.suffix.lower().endswith(("o", "rnx"))
    ]
    
    if not arquivos_o:
        print("Nenhum arquivo de observação encontrado.")
        return

    print(f"Iniciando PPP para {len(arquivos_o)} arquivos...")
    
    # Processamento sequencial: um arquivo por vez
    for obs in arquivos_o:
        resultado = processar_ppp_rtklib(
           obs,
           path_produtos,
            path_nav,
            path_config,
            path_rnx2rtkp,
            path_saida
        )
        print(resultado)

    print(f"\n🏁 Processamento finalizado. Verifique a pasta: {path_saida}")

if __name__ == "__main__":
    main()