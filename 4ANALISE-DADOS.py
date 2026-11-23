import os

os.environ.pop("LD_LIBRARY_PATH", None)

import math
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from pathlib import Path
import utils

# PARÂMETROS AJUSTÁVEIS
LIMITE_ERRO_METROS = 0.10
FRACAO_MELHORES = 0.05

LIMITE_NORTE = 100
LIMITE_LESTE = 100
LIMITE_ALTITUDE = 200

# Posição das colunas numa linha "OUTPUT" do gLAB 6.0 (split por espaços)
COL_ANO, COL_DOY, COL_SOD = 1, 2, 3
COL_NSAT = 7
COL_LAT, COL_LON, COL_ALT = 20, 21, 22
COL_SDN, COL_SDE, COL_SDU = 26, 27, 28
N_COLUNAS_MIN = 29


# LEITURA DO ARQUIVO .pos DO gLAB
def ler_pos_glab(arquivo):
    """Lê as linhas 'OUTPUT' de um arquivo de saída do gLAB e devolve um DataFrame."""
    linhas = []
    with open(arquivo, encoding="latin1", errors="ignore") as f:
        for l in f:
            if not l.startswith("OUTPUT"):
                continue
            c = l.split()
            if len(c) < N_COLUNAS_MIN:
                continue
            try:
                lat, lon, alt = float(c[COL_LAT]), float(c[COL_LON]), float(c[COL_ALT])
                if abs(lat) > 90 or abs(lon) > 180:
                    continue
                linhas.append({
                    'Ano': int(c[COL_ANO]), 'DoY': int(c[COL_DOY]), 'SoD': float(c[COL_SOD]),
                    'ns': int(c[COL_NSAT]),
                    'Lat': lat, 'Lon': lon, 'Height': alt,
                    'sdn': float(c[COL_SDN]), 'sde': float(c[COL_SDE]), 'sdu': float(c[COL_SDU]),
                })
            except ValueError:
                continue
    return pd.DataFrame(linhas)


# CONVERSÃO PARA mm (plano local N/E/U, com raios de curvatura do WGS84)
def calcular_variacao_milimetros(df):
    media_lat = df['Lat'].mean()
    media_lon = df['Lon'].mean()
    media_alt = df['Alt'].mean()

    a = 6378137.0
    e2 = 0.00669437999014
    phi = math.radians(media_lat)
    w = math.sqrt(1 - e2 * math.sin(phi) ** 2)
    raio_meridiano = a * (1 - e2) / w ** 3      # M
    raio_vertical = a / w                       # N

    df['dN (mm)'] = [math.radians(x - media_lat) * (raio_meridiano + media_alt) * 1000 for x in df['Lat']]
    df['dE (mm)'] = [math.radians(x - media_lon) * (raio_vertical + media_alt) * math.cos(phi) * 1000 for x in df['Lon']]
    df['dU (mm)'] = (df['Alt'] - media_alt) * 1000
    return df, media_lat, media_lon, media_alt


def processar_e_plotar(arquivos, constelacao, pasta_resultados):
    print(f"\n🔍 Processando {len(arquivos)} arquivos para a constelação: {constelacao}...")
    dados_diarios = []

    for arquivo in sorted(arquivos):
        try:
            df = ler_pos_glab(arquivo)
            if df.empty:
                print(f"      ⚠️ {arquivo.name}: sem linhas OUTPUT (o gLAB não gerou solução). Ignorado.")
                continue

            # incerteza 3D formal
            df['sd_3d'] = (df['sdn'] ** 2 + df['sde'] ** 2 + df['sdu'] ** 2) ** 0.5

            # 1) só épocas já convergidas
            df = df[df['sd_3d'] <= LIMITE_ERRO_METROS]
            if df.empty:
                print(f"      ⚠️ {arquivo.name}: Ignorado (nenhuma época com incerteza 3D <= {LIMITE_ERRO_METROS*100:.0f} cm)")
                continue

            # 2) os melhores X% entre elas
            df = df.sort_values('sd_3d')
            n_pontos = max(1, int(len(df) * FRACAO_MELHORES))
            melhores = df.head(n_pontos)

            ano, doy = int(df['Ano'].iloc[0]), int(df['DoY'].iloc[0])
            data_dia = pd.Timestamp(year=ano, month=1, day=1) + pd.Timedelta(days=doy - 1)

            dados_diarios.append({
                'Data': data_dia,
                'Lat': melhores['Lat'].median(),
                'Lon': melhores['Lon'].median(),
                'Alt': melhores['Height'].median(),
                'N_epocas_usadas': n_pontos,
                'Sat_medio': round(melhores['ns'].mean(), 1),
            })
        except Exception as e:
            print(f"⚠️ Erro ao processar {arquivo.name}: {e}")

    if not dados_diarios:
        print(f"❌ Nenhum dado válido pôde ser extraído para {constelacao}.")
        return

    df_resumo = pd.DataFrame(dados_diarios).sort_values('Data').reset_index(drop=True)
    df_resumo, m_lat, m_lon, m_alt = calcular_variacao_milimetros(df_resumo)

    if len(df_resumo) < 2:
        print(f"      ℹ️ Só {len(df_resumo)} dia com solução em {constelacao}: desvio padrão não pode ser calculado (NaN).")
    std_n, std_e, std_u = df_resumo['dN (mm)'].std(), df_resumo['dE (mm)'].std(), df_resumo['dU (mm)'].std()
    rmse_n = (df_resumo['dN (mm)'] ** 2).mean() ** 0.5
    rmse_e = (df_resumo['dE (mm)'] ** 2).mean() ** 0.5
    rmse_u = (df_resumo['dU (mm)'] ** 2).mean() ** 0.5

    print(f"🎯 Média {constelacao}: Lat {m_lat:.8f}°, Lon {m_lon:.8f}°, Alt {m_alt:.3f}m  ({len(df_resumo)} dias)")
    print("📊 ESTATÍSTICAS DA SÉRIE (Variação em milímetros):")
    for nome, d, s, r in [("NORTE", 'dN (mm)', std_n, rmse_n), ("LESTE", 'dE (mm)', std_e, rmse_e), ("ALT  ", 'dU (mm)', std_u, rmse_u)]:
        print(f"  {nome} -> Desvio Padrão: {s:.2f} | RMSE: {r:.2f} | Min: {df_resumo[d].min():.2f} | Máx: {df_resumo[d].max():.2f}")

    # Tabela
    pasta_tabelas = pasta_resultados / "Tabelas_Estatisticas"
    os.makedirs(pasta_tabelas, exist_ok=True)
    df_resumo.to_csv(pasta_tabelas / f"Dados_Serie_{constelacao}.csv", index=False, sep=';', decimal=',')

    # Gráficos
    plt.style.use('ggplot')
    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(12, 9), sharex=True)
    fig.suptitle(f'Série Temporal de Posição PPP - {constelacao}\nVariação Diária (Norte, Leste, Altitude)',
                 fontsize=16, fontweight='bold')
    estilo = {'marker': 'o', 'markersize': 5, 'linewidth': 1.5, 'alpha': 0.8}

    ax1.plot(df_resumo['Data'], df_resumo['dN (mm)'], color='tab:blue', **estilo)
    ax1.set_ylabel('Norte (mm)', fontweight='bold')
    ax1.set_ylim(-LIMITE_NORTE, LIMITE_NORTE)
    ax1.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)

    caixa_texto = (f"ESTATÍSTICAS (1σ)\n\n"
                   f"N: ± {std_n:.1f} mm\n"
                   f"E: ± {std_e:.1f} mm\n"
                   f"U: ± {std_u:.1f} mm\n\n"
                   f"MÉDIAS COORDS\n\n"
                   f"Lat: {m_lat:.8f}°\n"
                   f"Lon: {m_lon:.8f}°\n"
                   f"Alt: {m_alt:.3f}m")
    props = dict(boxstyle='round,pad=0.5', facecolor='white', alpha=0.9, edgecolor='gray')
    ax1.text(1.02, 0.5, caixa_texto, transform=ax1.transAxes, fontsize=11, fontweight='bold',
             verticalalignment='center', bbox=props)

    ax2.plot(df_resumo['Data'], df_resumo['dE (mm)'], color='tab:orange', **estilo)
    ax2.set_ylabel('Leste (mm)', fontweight='bold')
    ax2.set_ylim(-LIMITE_LESTE, LIMITE_LESTE)
    ax2.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)

    ax3.plot(df_resumo['Data'], df_resumo['dU (mm)'], color='tab:green', **estilo)
    ax3.set_ylabel('Altitude (mm)', fontweight='bold')
    ax3.set_xlabel('Data da Observação', fontweight='bold')
    ax3.set_ylim(-LIMITE_ALTITUDE, LIMITE_ALTITUDE)
    ax3.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)

    span_dias = (df_resumo['Data'].max() - df_resumo['Data'].min()).days
    if 1 <= span_dias <= 120:
        ax3.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, span_dias // 10)))
    elif span_dias == 0:   # um único dia: centraliza o ponto
        d0 = df_resumo['Data'].iloc[0]
        ax3.set_xlim(d0 - pd.Timedelta(days=1), d0 + pd.Timedelta(days=1))
        ax3.xaxis.set_major_locator(mdates.DayLocator())
    ax3.xaxis.set_major_formatter(mdates.DateFormatter('%d/%b/%Y'))
    plt.gcf().autofmt_xdate()

    pasta_graficos = pasta_resultados / "Graficos_Serie_Temporal"
    os.makedirs(pasta_graficos, exist_ok=True)
    caminho_grafico = pasta_graficos / f"Serie_Temporal_{constelacao}.png"
    plt.tight_layout()
    plt.savefig(caminho_grafico, dpi=300, bbox_inches='tight')
    print(f"✅ Gráfico salvo: {caminho_grafico.name}")
    plt.close()


def main():
    print("📈 GERADOR DE SÉRIE TEMPORAL (saída do gLAB) - GPS / GLONASS / GPS+GLONASS")

    caminho_salvo = utils.carregar_estado("pasta_rinex_pronta")
    if not caminho_salvo:
        print("❌ Memória não encontrada. Execute o script 1 primeiro.")
        return

    path_rinex_obs = Path(caminho_salvo)
    if path_rinex_obs.name in ["GPS", "GLONASS", "GPS_GLONASS"]:
        path_rinex_obs = path_rinex_obs.parent

    pasta_resultados = path_rinex_obs.parent / "RESULTADOS_PPP"
    arquivos_pos = [f for f in pasta_resultados.glob("*.pos") if "events" not in f.name]

    if not arquivos_pos:
        print(f"❌ Nenhum arquivo .pos encontrado na pasta: {pasta_resultados}")
        return

    # Ordem importa: GPS_GLONASS_ começa com "GPS_", então é separado primeiro
    arq_gps_glo = [f for f in arquivos_pos if f.name.startswith("GPS_GLONASS_")]
    arq_gps = [f for f in arquivos_pos if f.name.startswith("GPS_") and not f.name.startswith("GPS_GLONASS_")]
    arq_glo = [f for f in arquivos_pos if f.name.startswith("GLONASS_")]

    for nome_const, lista_arq in [("GPS", arq_gps), ("GLONASS", arq_glo), ("GPS_GLONASS", arq_gps_glo)]:
        if not lista_arq:
            continue
        try:
            processar_e_plotar(lista_arq, nome_const, pasta_resultados)
        except Exception as e:
            print(f"💥 Falha ao processar/plotar {nome_const}: {e}")

    print("\n🎉 Séries temporais geodésicas concluídas!")


if __name__ == "__main__":
    main()