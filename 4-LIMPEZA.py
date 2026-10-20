import argparse
import os
import re
import shutil
import stat
import sys
from pathlib import Path

import utils

# =============================================================================
# O QUE É REMOVIDO (somente dentro da pasta do mês atual: .../MES_AA/)
# =============================================================================
# Pastas criadas pelos scripts 1 e 2 (nomes exatamente como nos scripts)
PASTAS_REMOVER = [
    "1 - Dados tipos .d",                                      # script 1 
    "1 - Dados tipo .d",                                       # variação de grafia, por segurança
    "1.1 - Navegacao Broadcast",                               # script 1 (BRDC)
    "2 - Dados separados por satélite (Prontos para RTKLIB)",  # script 1 (GPS, GLONASS, GPS_GLONASS)
    "IGS_PRODUCTS",                                            # script 2 (SP3 / CLK)
    "TEMP_ZIPS",                                               # script 1 (só sobram se o script 1 foi interrompido)
    "TEMP_EXTRAIDOS",                                          # idem
]

# Dentro de RESULTADOS_PPP (script 3): arquivos de debug do rnx2rtkp (-x 3 / -y 3)
# Ex.: ARQUIVO.pos.stat, ARQUIVO.pos.trace
SUFIXOS_RESULTADOS = (".stat", ".trace")
# Configs temporárias do script 3 (ficam para trás se o RTKLIB falhar)
PADRAO_CONF_TEMP = "_tmp_*.conf"

PADRAO_PASTA_MES = re.compile(r"^[A-Za-z]{3}_\d{2}$")  # ex.: JAN_24
SUBPASTAS_CONSTELACAO = ["GPS", "GLONASS", "GPS_GLONASS"]


# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================
def tamanho(caminho):
    """Tamanho total (bytes) de um arquivo ou pasta."""
    caminho = Path(caminho)
    if caminho.is_file():
        return caminho.stat().st_size
    total = 0
    for raiz, _, arquivos in os.walk(caminho):
        for a in arquivos:
            try:
                total += (Path(raiz) / a).stat().st_size
            except OSError:
                pass
    return total


def _tornar_gravavel_e_repetir(func, caminho, *_):
    """Handler do rmtree: arquivos somente-leitura (comum no Windows)."""
    try:
        os.chmod(caminho, stat.S_IWRITE)
        func(caminho)
    except Exception:
        pass


def remover(caminho, simular):
    """Remove arquivo ou pasta. Retorna os bytes liberados (ou que seriam liberados)."""
    caminho = Path(caminho)
    tam = tamanho(caminho) if not caminho.is_symlink() else 0
    if simular:
        return tam
    if caminho.is_symlink() or caminho.is_file():
        try:
            caminho.unlink()
        except PermissionError:
            os.chmod(caminho, stat.S_IWRITE)
            caminho.unlink()
    else:
        if sys.version_info >= (3, 12):
            shutil.rmtree(caminho, onexc=_tornar_gravavel_e_repetir)
        else:
            shutil.rmtree(caminho, onerror=_tornar_gravavel_e_repetir)
    return tam


def resolver_pasta_mes(pasta_arg=None):
    """
    Descobre a pasta do mês (MES_AA). Usa o argumento, se houver; caso contrário o
    estado salvo pelo script 1 ('pasta_rinex_pronta' = .../MES_AA/2 - Dados separados...).
    """
    if pasta_arg:
        return Path(pasta_arg)

    caminho_salvo = utils.carregar_estado("pasta_rinex_pronta")
    if not caminho_salvo:
        return None

    pasta_rinex = Path(caminho_salvo)
    if pasta_rinex.name in SUBPASTAS_CONSTELACAO:
        pasta_rinex = pasta_rinex.parent
    return pasta_rinex.parent


def validar_pasta_mes(pasta_mes, forcar):
    """Trava de segurança: só limpa se for mesmo uma pasta de mês já analisada."""
    if not pasta_mes.is_dir():
        print(f"❌ Pasta do mês não encontrada: {pasta_mes}")
        return False

    if not PADRAO_PASTA_MES.match(pasta_mes.name):
        print(f"❌ '{pasta_mes.name}' não parece uma pasta de mês (esperado ex.: JAN_24). Abortando.")
        return False

    pasta_res = pasta_mes / "RESULTADOS_PPP"
    if not pasta_res.is_dir():
        print(f"❌ RESULTADOS_PPP não existe em {pasta_mes}. Nada a limpar com segurança.")
        return False

    if forcar:
        return True

    tem_pos = any(f for f in pasta_res.glob("*.pos"))
    pasta_tab = pasta_res / "Tabelas_Estatisticas"
    tem_csv = pasta_tab.is_dir() and any(pasta_tab.glob("Dados_Serie_*.csv"))

    if not tem_pos:
        print("❌ Nenhum .pos em RESULTADOS_PPP: o script 3 parece não ter terminado. Abortando.")
        return False
    if not tem_csv:
        print("❌ Nenhum Dados_Serie_*.csv em Tabelas_Estatisticas: o script 4 parece não ter rodado.")
        print("   Os dados de entrada NÃO foram apagados. (Use --forcar para ignorar esta trava.)")
        return False
    return True


# =============================================================================
# LIMPEZA
# =============================================================================
def limpar_mes(pasta_mes=None, simular=False, forcar=False):
    pasta_mes = resolver_pasta_mes(pasta_mes)
    if pasta_mes is None:
        print("❌ Não foi possível determinar a pasta do mês (estado não encontrado). "
              "Informe o caminho com --pasta.")
        return False

    print(f"🧹 LIMPEZA DO MÊS: {pasta_mes}" + ("  [SIMULAÇÃO - nada será apagado]" if simular else ""))

    if not validar_pasta_mes(pasta_mes, forcar):
        return False

    total = 0

    # 1. Pastas de processamento
    for nome in PASTAS_REMOVER:
        alvo = pasta_mes / nome
        if alvo.exists():
            try:
                tam = remover(alvo, simular)
                total += tam
                print(f"   🗑️  Pasta: {nome}  ({tam / 1024**2:.1f} MB)")
            except Exception as e:
                print(f"   ❌ Falha ao remover {nome}: {e}")

    # 2. Arquivos de debug e configs temporárias em RESULTADOS_PPP (somente nível raiz)
    pasta_res = pasta_mes / "RESULTADOS_PPP"
    for arq in sorted(pasta_res.iterdir()):
        if not arq.is_file():
            continue
        nome_l = arq.name.lower()
        eh_debug = nome_l.endswith(SUFIXOS_RESULTADOS)
        eh_conf_tmp = arq.match(PADRAO_CONF_TEMP)
        if eh_debug or eh_conf_tmp:
            try:
                tam = remover(arq, simular)
                total += tam
                print(f"   🗑️  Arquivo: RESULTADOS_PPP/{arq.name}  ({tam / 1024**2:.1f} MB)")
            except Exception as e:
                print(f"   ❌ Falha ao remover {arq.name}: {e}")

    verbo = "seriam liberados" if simular else "liberados"
    print(f"\n✅ Limpeza concluída: {total / 1024**2:.1f} MB {verbo}.")
    print("   Mantidos: .pos, Tabelas_Estatisticas (.csv), Graficos_Serie_Temporal e demais arquivos.")
    return True


def main():
    parser = argparse.ArgumentParser(description="Limpeza dos arquivos intermediários do mês processado.")
    parser.add_argument("--pasta", help="Pasta do mês (ex.: D:\\GNSS\\2024\\JAN_24). "
                                        "Se omitida, usa o estado salvo pelo script 1.")
    parser.add_argument("--simular", action="store_true", help="Apenas lista o que seria removido.")
    parser.add_argument("--forcar", action="store_true",
                        help="Ignora a verificação de que os scripts 3 e 4 já geraram .pos e .csv.")
    args = parser.parse_args()

    limpar_mes(args.pasta, simular=args.simular, forcar=args.forcar)


if __name__ == "__main__":
    main()