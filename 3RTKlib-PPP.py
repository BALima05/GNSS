import os
import subprocess
from pathlib import Path
import concurrent.futures

def processar_ppp_rtklib(arquivo_obs, pasta_produtos, config_file, rnx2rtkp_path, pasta_saida):
    try:
        arquivo_obs = Path(arquivo_obs)
        pasta_produtos = Path(pasta_produtos)
        pasta_saida = Path(pasta_saida)
        
        # Nome do arquivo de saída
        arquivo_pos = pasta_saida / arquivo_obs.with_suffix('.pos').name
        
        # --- CORREÇÃO 1: Buscar também arquivos de Navegação (.n, .p, .nav) ---
        # Tenta achar arquivos de navegação na pasta de produtos OU na pasta do arquivo .o
        nav_files = list(pasta_produtos.glob("*.[0-9][0-9]n")) + \
                    list(pasta_produtos.glob("*.[0-9][0-9]p")) + \
                    list(pasta_produtos.glob("*.nav")) + \
                    list(arquivo_obs.parent.glob(f"*{arquivo_obs.suffix[-3:-1]}n")) # Ex: procura .24n se o obs for .24o

        arquivos_sp3 = list(pasta_produtos.glob("*.sp3")) + list(pasta_produtos.glob("*.eph"))
        arquivos_clk = list(pasta_produtos.glob("*.clk"))
        
        if not arquivos_sp3:
            return f"⚠️ Pulei {arquivo_obs.name}: Faltam arquivos .sp3 (Orbitas)."
            
        # Monta o comando
        cmd = [
            str(rnx2rtkp_path),
            '-k', str(config_file),
            '-o', str(arquivo_pos),
            str(arquivo_obs)
        ]
        
        # Adiciona produtos e navegação ao comando
        cmd.extend([str(p) for p in arquivos_sp3])
        cmd.extend([str(p) for p in arquivos_clk])
        cmd.extend([str(p) for p in nav_files]) # Adiciona navegação

        # Executa capturando TUDO
        result = subprocess.run(cmd, capture_output=True, text=True)

        # --- CORREÇÃO 2: Verificar se o arquivo EXISTE e tem CONTEÚDO ---
        if arquivo_pos.exists() and arquivo_pos.stat().st_size > 0:
            return f"✅ PPP Sucesso: {arquivo_pos.name} (Tamanho: {arquivo_pos.stat().st_size/1024:.1f} KB)"
        else:
            # Se o arquivo não foi criado, mostra o erro que o RTKLIB cuspiu
            return (f"❌ Falha {arquivo_obs.name} (Arquivo vazio ou não criado).\n"
                    f"   Log RTKLIB: {result.stderr}\n"
                    f"   Output: {result.stdout}")

    except Exception as e:
        return f"💥 Erro de execução: {e}"

def main():
    print("🌍 AUTOMÇÃO DE PPP COM RTKLIB (Python Wrapper)")
    
    # --- CONFIGURAÇÕES ---
    # Caminho para o executável rnx2rtkp.exe
    path_rnx2rtkp = r"C:\Users\berna\Downloads\RTKLIB_EX_2.5.0\RTKLIB_EX_2.5.0\rnx2rtkp.exe"
    
    # Pasta onde estão seus arquivos RINEX .o (gerados no script anterior)
    path_rinex_obs = input("Pasta com arquivos RINEX (.o): ").strip().strip('"')
    
    # Pasta onde você salvou os arquivos .sp3 e .clk baixados do IGS
    path_produtos = input("Pasta com produtos IGS (.sp3/.clk): ").strip().strip('"')
    
    # Arquivo de configuração .conf
    path_config = r"C:\Users\berna\OneDrive\Documentos\PUB IC\GNSS\ppp-static.conf"
    
    # Pasta para salvar os resultados
    path_saida = os.path.join(os.path.dirname(path_rinex_obs), "RESULTADOS_PPP")
    os.makedirs(path_saida, exist_ok=True)
    
    # --- PROCESSAMENTO ---
    path_rinex_obs = Path(path_rinex_obs)
    arquivos_o = list(path_rinex_obs.glob("*.[0-9][0-9]o"))
    
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