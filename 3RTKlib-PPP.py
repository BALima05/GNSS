import os
import subprocess
from pathlib import Path
import concurrent.futures
import config
import utils

def processar_ppp_rtklib(arquivo_obs, pasta_produtos, pasta_nav, config_file, rnx2rtkp_path, pasta_saida):
    try:
        arquivo_obs = Path(arquivo_obs)
        pasta_produtos = Path(pasta_produtos)
        pasta_nav = Path(pasta_nav)
        pasta_saida = Path(pasta_saida)
        
        # Nome do arquivo de saída
        arquivo_pos = pasta_saida / arquivo_obs.with_suffix('.pos').name
        
        # Obtenção dos arquivos nav e produtos
        nav_files = list(pasta_nav.glob("*.[0-9][0-9]n")) + \
                    list(pasta_nav.glob("*.[0-9][0-9]p")) + \
                    list(pasta_nav.glob("*.[0-9][0-9]g")) + \
                    list(pasta_nav.glob("*.nav"))

        arquivos_sp3 = list(pasta_produtos.glob("*.[sS][pP]3")) + list(pasta_produtos.glob("*.eph"))
        arquivos_clk = list(pasta_produtos.glob("*.[cC][lL][kK]"))
        
        if not arquivos_sp3:
            return f"⚠️ Pulei {arquivo_obs.name}: Faltam arquivos .sp3 (Orbitas)."
        if not arquivos_clk:
            return f"⚠️ Pulei {arquivo_obs.name}: Faltam arquivos .clk (Relógios)."
        if not nav_files:
            return f"⚠️ Pulei {arquivo_obs.name}: Faltam arquivos de navegação (.n, .p, .g, .nav)."
            
        # Monta o comando
        cmd = [
            str(rnx2rtkp_path),
            '-k', str(config_file),
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

        # Executa capturando TUDO
        result = subprocess.run(cmd, capture_output=True, text=True)

        # --- Verificar se o arquivo EXISTE e tem CONTEÚDO ---
        if arquivo_pos.exists() and arquivo_pos.stat().st_size > 0:
            return f"✅ PPP Sucesso: {arquivo_pos.name} (Tamanho: {arquivo_pos.stat().st_size/1024:.1f} KB)"
        else:
            # Se o arquivo não foi criado, mostra o erro que o RTKLIB cuspiu
            return (f"❌ Falha {arquivo_obs.name} (Arquivo vazio ou não criado).\n"
                    f"   Log RTKLIB: {result.stderr}\n"
                    f"   Output: {result.stdout}")

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
    arquivos_o = list(path_rinex_obs.rglob("*.*o"))
    
    if not arquivos_o:
        print("Nenhum arquivo de observação encontrado.")
        return

    print(f"Iniciando PPP para {len(arquivos_o)} arquivos...")
    
    # Processamento Paralelo (PPP consome CPU, cuidado com muitos núcleos)
    with concurrent.futures.ProcessPoolExecutor() as executor:
        tarefas = {
            executor.submit(
                processar_ppp_rtklib, 
                obs, 
                path_produtos,
                path_nav,
                path_config, 
                path_rnx2rtkp, 
                path_saida
            ): obs for obs in arquivos_o
        }
        
        for futuro in concurrent.futures.as_completed(tarefas):
            print(futuro.result())

    print(f"\n🏁 Processamento finalizado. Verifique a pasta: {path_saida}")

if __name__ == "__main__":
    main()