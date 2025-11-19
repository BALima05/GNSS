import os
import subprocess
from pathlib import Path
import concurrent.futures

def processar_ppp_rtklib(arquivo_obs, pasta_produtos, config_file, rnx2rtkp_path, pasta_saida):
    """
    Executa o rnx2rtkp (RTKLIB) em modo PPP para um arquivo de observação.
    """
    try:
        arquivo_obs = Path(arquivo_obs)
        pasta_produtos = Path(pasta_produtos)
        pasta_saida = Path(pasta_saida)
        
        # Define o nome do arquivo de saída (.pos)
        arquivo_pos = pasta_saida / arquivo_obs.with_suffix('.pos').name
        
        # Encontra arquivos .sp3 (órbitas) e .clk (relógios) na pasta de produtos
        # DICA: O RTKLIB é inteligente. Se você passar vários arquivos .sp3/.clk, 
        # ele usa apenas os que correspondem ao horário do arquivo .o.
        arquivos_sp3 = list(pasta_produtos.glob("*.sp3")) + list(pasta_produtos.glob("*.eph"))
        arquivos_clk = list(pasta_produtos.glob("*.clk"))
        
        if not arquivos_sp3 or not arquivos_clk:
            return f"⚠️ Pulei {arquivo_obs.name}: Faltam arquivos .sp3 ou .clk na pasta de produtos."

        # Monta o comando do RTKLIB
        # rnx2rtkp -k config.conf -o saida.pos obs.o orbita.sp3 relogio.clk
        cmd = [
            str(rnx2rtkp_path),
            '-k', str(config_file),
            '-o', str(arquivo_pos),
            str(arquivo_obs)
        ]
        
        # Adiciona todos os arquivos de produto ao comando
        cmd.extend([str(p) for p in arquivos_sp3])
        cmd.extend([str(p) for p in arquivos_clk])

        # Executa o comando (sem shell=True para segurança e melhor manuseio de lista)
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode == 0:
            return f"✅ PPP Sucesso: {arquivo_pos.name}"
        else:
            return f"❌ Erro no RTKLIB para {arquivo_obs.name}:\n{result.stderr}"

    except Exception as e:
        return f"💥 Erro de execução: {e}"

def main():
    print("🌍 AUTOMÇÃO DE PPP COM RTKLIB (Python Wrapper)")
    
    # --- CONFIGURAÇÕES ---
    # Caminho para o executável rnx2rtkp.exe
    path_rnx2rtkp = input("Caminho do rnx2rtkp.exe: ").strip().strip('"')
    
    # Pasta onde estão seus arquivos RINEX .o (gerados no script anterior)
    path_rinex_obs = input("Pasta com arquivos RINEX (.o): ").strip().strip('"')
    
    # Pasta onde você salvou os arquivos .sp3 e .clk baixados do IGS
    path_produtos = input("Pasta com produtos IGS (.sp3/.clk): ").strip().strip('"')
    
    # Arquivo de configuração .conf
    path_config = input("Caminho do arquivo ppp_static.conf: ").strip().strip('"')
    
    # Pasta para salvar os resultados
    path_saida = os.path.join(os.path.dirname(path_rinex_obs), "RESULTADOS_PPP")
    os.makedirs(path_saida, exist_ok=True)
    
    # --- PROCESSAMENTO ---
    path_rinex_obs = Path(path_rinex_obs)
    arquivos_o = list(path_rinex_obs.glob("*.o")) # ou *.XXo se não tiver renomeado
    
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