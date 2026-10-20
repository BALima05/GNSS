import os

os.environ.pop("LD_LIBRARY_PATH", None)

import re
import sys
import math
from pathlib import Path

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# =============================================================================
# CONFIGURAÇÕES
# =============================================================================
# Pasta do ANO (que contém JAN_24, FEV_24, ...). Pode ser passada como argumento:
#   python 5ANALISE-ANUAL.py "D:\GNSS\2024"
# Se não for passada, tenta usar config.PASTA_BASE; ajuste aqui se preferir.
PASTA_ANO_PADRAO = None

CONSTELACOES = ["GPS", "GLONASS", "GPS_GLONASS"]

# Limites dos eixos dos gráficos (mm) - mesmos do script 4
LIMITE_NORTE = 100
LIMITE_LESTE = 100
LIMITE_ALTITUDE = 200

PADRAO_PASTA_MES = re.compile(r"^[A-Za-z]{3}_\d{2}$")  # ex.: JAN_24


# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================
def resolver_pasta_ano():
    """Define a pasta do ano: argumento > constante > config.PASTA_BASE."""
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    if PASTA_ANO_PADRAO:
        return Path(PASTA_ANO_PADRAO)
    try:
        import config
        return Path(config.PASTA_BASE)
    except Exception:
        return None


def calcular_variacao_milimetros(df):
    """Mesma lógica do script 4, agora com a média GLOBAL DO ANO."""
    media_lat = df['Lat'].mean()
    media_lon = df['Lon'].mean()
    media_alt = df['Alt'].mean()

    fator_lat = 111320.0
    fator_lon = 111320.0 * math.cos(math.radians(media_lat))

    df['dN (mm)'] = (df['Lat'] - media_lat) * fator_lat * 1000
    df['dE (mm)'] = (df['Lon'] - media_lon) * fator_lon * 1000
    df['dU (mm)'] = (df['Alt'] - media_alt) * 1000
    return df, media_lat, media_lon, media_alt


def calcular_estatisticas(df, rotulo):
    """
    Estatísticas (desvio padrão, RMSE, mín, máx) de dN/dE/dU de um DataFrame,
    em relação à média do próprio DataFrame. Retorna um dict (uma linha da tabela).
    """
    # Recalcula as variações em relação à média do próprio conjunto
    # (para os meses, equivale ao que o script 4 fez em cada mês)
    d = df[['Lat', 'Lon', 'Alt']].copy()
    d, m_lat, m_lon, m_alt = calcular_variacao_milimetros(d)

    linha = {
        'Periodo': rotulo,
        'N_dias': len(d),
        'Lat_media (graus)': m_lat,
        'Lon_media (graus)': m_lon,
        'Alt_media (m)': m_alt,
    }
    for eixo, col in [('N', 'dN (mm)'), ('E', 'dE (mm)'), ('U', 'dU (mm)')]:
        serie = d[col]
        linha[f'Std_{eixo} (mm)'] = serie.std() if len(serie) > 1 else float('nan')
        linha[f'RMSE_{eixo} (mm)'] = (serie ** 2).mean() ** 0.5
        linha[f'Min_{eixo} (mm)'] = serie.min()
        linha[f'Max_{eixo} (mm)'] = serie.max()
    return linha


def carregar_series_mensais(pasta_ano, constelacao):
    """
    Lê os CSVs mensais 'Dados_Serie_{constelacao}.csv' gerados pelo script 4
    de cada pasta MES_ANO e concatena tudo em um único DataFrame.
    """
    pastas_mes = sorted(
        d for d in pasta_ano.iterdir()
        if d.is_dir() and PADRAO_PASTA_MES.match(d.name)
    )

    quadros = []
    for pasta_mes in pastas_mes:
        csv_mes = pasta_mes / "RESULTADOS_PPP" / "Tabelas_Estatisticas" / f"Dados_Serie_{constelacao}.csv"
        if not csv_mes.exists():
            print(f"   ⚠️ {pasta_mes.name}: sem {csv_mes.name} (mês ignorado para {constelacao})")
            continue
        try:
            df = pd.read_csv(csv_mes, sep=';', decimal=',', parse_dates=['Data'])
            df = df[['Data', 'Lat', 'Lon', 'Alt']].dropna()
            df['Mes_Pasta'] = pasta_mes.name.upper()
            quadros.append(df)
            print(f"   📄 {pasta_mes.name}: {len(df)} dias")
        except Exception as e:
            print(f"   ⚠️ Erro lendo {csv_mes}: {e}")

    if not quadros:
        return None

    df_ano = pd.concat(quadros, ignore_index=True)
    df_ano = df_ano.drop_duplicates(subset='Data').sort_values('Data').reset_index(drop=True)
    return df_ano


def salvar_csv_anual(df_resumo, tabela_stats, caminho):
    """
    Salva UM csv com: (1) série diária do ano e (2) logo abaixo, a tabela de
    estatísticas (ANUAL + cada mês). Separador ';' e decimal ',' como no script 4.
    """
    texto = "SERIE DIARIA\n"
    texto += df_resumo.to_csv(index=False, sep=';', decimal=',')
    texto += "\nESTATISTICAS (ANUAL e MENSAIS)\n"
    texto += tabela_stats.to_csv(index=False, sep=';', decimal=',')
    with open(caminho, 'w', encoding='utf-8-sig', newline='') as f:
        f.write(texto)


def plotar_serie_anual(df_resumo, constelacao, ano_txt, stats_ano, pasta_saida):
    """Gráfico no mesmo formato do script 4 (3 painéis N/E/U + caixa de estatísticas)."""
    plt.style.use('ggplot')
    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(14, 9), sharex=True)
    fig.suptitle(
        f'Série Temporal Anual de Posição PPP - {constelacao} ({ano_txt})\n'
        f'Variação Diária (Norte, Leste, Altitude)',
        fontsize=16, fontweight='bold')

    estilo = {'marker': 'o', 'markersize': 4, 'linewidth': 1.2, 'alpha': 0.8}

    ax1.plot(df_resumo['Data'], df_resumo['dN (mm)'], color='tab:blue', **estilo)
    ax1.set_ylabel('Norte (mm)', fontweight='bold')
    ax1.set_ylim(-LIMITE_NORTE, LIMITE_NORTE)

    ax2.plot(df_resumo['Data'], df_resumo['dE (mm)'], color='tab:orange', **estilo)
    ax2.set_ylabel('Leste (mm)', fontweight='bold')
    ax2.set_ylim(-LIMITE_LESTE, LIMITE_LESTE)

    ax3.plot(df_resumo['Data'], df_resumo['dU (mm)'], color='tab:green', **estilo)
    ax3.set_ylabel('Altitude (mm)', fontweight='bold')
    ax3.set_xlabel('Data da Observação', fontweight='bold')
    ax3.set_ylim(-LIMITE_ALTITUDE, LIMITE_ALTITUDE)

    for ax in (ax1, ax2, ax3):
        ax.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)

    caixa_texto = (f"ESTATÍSTICAS ANUAIS (1σ)\n\n"
                   f"N: ± {stats_ano['Std_N (mm)']:.1f} mm\n"
                   f"E: ± {stats_ano['Std_E (mm)']:.1f} mm\n"
                   f"U: ± {stats_ano['Std_U (mm)']:.1f} mm\n\n"
                   f"MÉDIAS COORDS\n\n"
                   f"Lat: {stats_ano['Lat_media (graus)']:.8f}°\n"
                   f"Lon: {stats_ano['Lon_media (graus)']:.8f}°\n"
                   f"Alt: {stats_ano['Alt_media (m)']:.3f}m\n\n"
                   f"Dias: {stats_ano['N_dias']}")
    props = dict(boxstyle='round,pad=0.5', facecolor='white', alpha=0.9, edgecolor='gray')
    ax1.text(1.02, 0.5, caixa_texto, transform=ax1.transAxes, fontsize=11, fontweight='bold',
             verticalalignment='center', bbox=props)

    ax3.xaxis.set_major_locator(mdates.MonthLocator())
    ax3.xaxis.set_major_formatter(mdates.DateFormatter('%b/%Y'))
    plt.gcf().autofmt_xdate()

    caminho = pasta_saida / f"Serie_Temporal_Anual_{constelacao}.png"
    plt.tight_layout()
    plt.savefig(caminho, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"✅ Gráfico salvo: {caminho.name}")


def processar_constelacao_anual(pasta_ano, constelacao, pasta_anual):
    print(f"\n🔍 Consolidando ano para a constelação: {constelacao}")
    df = carregar_series_mensais(pasta_ano, constelacao)
    if df is None or df.empty:
        print(f"❌ Nenhum dado mensal encontrado para {constelacao}.")
        return

    # Ano de referência (pelo conteúdo dos dados)
    ano_completo = int(df['Data'].dt.year.mode().iloc[0])
    ano_txt = str(ano_completo)

    # Variações em relação à média ANUAL (eixos do gráfico e colunas dN/dE/dU do CSV)
    df_resumo, m_lat, m_lon, m_alt = calcular_variacao_milimetros(df.copy())

    # Estatísticas: linha anual + uma linha por mês (cada mês vs. sua própria média,
    # exatamente como o script 4 fez; os dados são os mesmos dos CSVs mensais)
    stats_ano = calcular_estatisticas(df_resumo, f"ANUAL {ano_txt}")
    linhas = [stats_ano]
    df_resumo['Mes'] = df_resumo['Data'].dt.month
    for mes, grupo in df_resumo.groupby('Mes'):
        rotulo = grupo['Mes_Pasta'].iloc[0]
        linhas.append(calcular_estatisticas(grupo, rotulo))
    tabela_stats = pd.DataFrame(linhas)

    # Console
    print(f"🎯 Média anual {constelacao}: Lat {m_lat:.8f}°, Lon {m_lon:.8f}°, Alt {m_alt:.3f}m")
    print("📊 ESTATÍSTICAS ANUAIS (Variação em milímetros):")
    for eixo, nome in [('N', 'NORTE'), ('E', 'LESTE'), ('U', 'ALT  ')]:
        print(f"  {nome} -> Desvio Padrão: {stats_ano[f'Std_{eixo} (mm)']:.2f} | "
              f"RMSE: {stats_ano[f'RMSE_{eixo} (mm)']:.2f} | "
              f"Min: {stats_ano[f'Min_{eixo} (mm)']:.2f} | "
              f"Máx: {stats_ano[f'Max_{eixo} (mm)']:.2f}")

    # Salvar CSVs
    df_saida = df_resumo[['Data', 'Mes_Pasta', 'Lat', 'Lon', 'Alt', 'dN (mm)', 'dE (mm)', 'dU (mm)']]
    df_saida = df_saida.rename(columns={'Mes_Pasta': 'Mes'})

    csv_anual = pasta_anual / f"Dados_Serie_Anual_{constelacao}.csv"
    salvar_csv_anual(df_saida, tabela_stats, csv_anual)
    print(f"✅ CSV anual salvo: {csv_anual.name}")

    csv_stats = pasta_anual / f"Estatisticas_Mensais_Anual_{constelacao}.csv"
    tabela_stats.to_csv(csv_stats, index=False, sep=';', decimal=',', encoding='utf-8-sig')
    print(f"✅ Tabela de estatísticas salva: {csv_stats.name}")

    # Gráfico
    plotar_serie_anual(df_resumo, constelacao, ano_txt, stats_ano, pasta_anual)


def main():
    print("📈 ANÁLISE ESTATÍSTICA ANUAL - SÉRIE TEMPORAL PPP")

    pasta_ano = resolver_pasta_ano()
    if pasta_ano is None or not pasta_ano.is_dir():
        print(f"❌ Pasta do ano inválida: {pasta_ano}")
        print("   Informe o caminho como argumento ou ajuste PASTA_ANO_PADRAO no script.")
        return
    print(f"📂 Pasta do ano: {pasta_ano}")

    # Nome da pasta de saída: ANUAL_24 (usa os 2 últimos dígitos do nome da pasta do ano;
    # se o nome não for numérico, usa o ano encontrado nos nomes MES_AA)
    m = re.search(r"(\d{4})$", pasta_ano.name)
    if m:
        sufixo = m.group(1)[2:]
    else:
        sufixos = sorted({d.name[-2:] for d in pasta_ano.iterdir()
                          if d.is_dir() and PADRAO_PASTA_MES.match(d.name)})
        if not sufixos:
            print("❌ Nenhuma pasta MES_AA encontrada (ex.: JAN_24).")
            return
        sufixo = sufixos[-1]

    pasta_anual = pasta_ano / f"ANUAL_{sufixo}"
    os.makedirs(pasta_anual, exist_ok=True)

    for const in CONSTELACOES:
        try:
            processar_constelacao_anual(pasta_ano, const, pasta_anual)
        except Exception as e:
            print(f"💥 Falha ao processar {const}: {e}")

    print(f"\n🎉 Análise anual concluída! Resultados em: {pasta_anual}")


if __name__ == "__main__":
    main()