import subprocess
import time
import sys

# Lista dos scripts na ordem exata de execução
scripts = [
    "1IBGE-RBMC.py",
    "2BAIXAR-PRODUTOS.py",
    "3RTKlib-PPP.py",
    "4ANALISE-DADOS.py"
]

print("🚀 INICIANDO PROCESSAMENTO EM CADEIA")
print("===================================")

t_inicio = time.time()

for script in scripts:
    print(f"\n▶️ Executando: {script}...")
    
    # Chama o script como no terminal
    # 'python' chama o interpretador, 'script' é o arquivo
    processo = subprocess.run([sys.executable, script])
    
    # Verifica se o script terminou com erro (código diferente de 0)
    if processo.returncode != 0:
        print(f"\n❌ ERRO CRÍTICO: O script {script} falhou.")
        print("A execução em cadeia foi interrompida.")
        break
    else:
        print(f"✅ {script} finalizado com sucesso.")

t_total = (time.time() - t_inicio) / 60
print("\n===================================")
print(f"🏁 TUDO PRONTO! Tempo total: {t_total:.1f} minutos.")