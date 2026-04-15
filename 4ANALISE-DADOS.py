import os
import math
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from pathlib import Path
import utils

def calcular_variacao_milimetros(df):
    """
    Calcula o centro (média global) e converte a diferença de cada dia
    para o plano topocêntrico local (Norte, Leste, UP) em milímetros.
    """
    media_lat = df['Lat'].mean()
    media_lon = df['Lon'].mean()
    media_alt = df['Alt'].mean()

    # Constantes para conversão (1 grau ~ 111.32 km)
    fator_lat = 111320.0 
    fator_lon = 111320.0 * math.cos(math.radians(media_lat))

    # Diferença em relação à média convertida para Milímetros (* 1000)
    df['dN (mm)'] = (df['Lat'] - media_lat) * fator_lat * 1000
    df['dE (mm)'] = (df['Lon'] - media_lon) * fator_lon * 1000
    df['dU (mm)'] = (df['Alt'] - media_alt) * 1000

    return df, media_lat, media_lon, media_alt

def processar_e_plotar(arquivos, constelacao, pasta_resultados):
    """
    Lê os arquivos de uma constelação específica, extrai a mediana diária
    e gera um gráfico em PNG.
    """
    print(f"\n🔍 Processando {len(arquivos)} arquivos para a constelação: {constelacao}...")
    
    colunas_padrao = ['Date', 'Time', 'Lat', 'Lon', 'Height', 'Q', 'ns', 'sdn', 'sde', 'sdu', 'sdne', 'sdeu', 'sdun', 'age', 'ratio']
    dados_diarios = []

    for arquivo in arquivos:
        try:
            df_temp = pd.read_csv(arquivo, comment='%', sep=r'\s+', names=colunas_padrao, encoding='latin1')
            
            if df_temp.empty:
                continue

            data_dia = pd.to_datetime(df_temp['Date'].iloc[0])
            lat_mediana = df_temp['Lat'].median()
            lon_mediana = df_temp['Lon'].median()
            alt_mediana = df_temp['Height'].median()

            dados_diarios.append({
                'Data': data_dia,
                'Lat': lat_mediana,
                'Lon': lon_mediana,
                'Alt': alt_mediana
            })
        except Exception as e:
            print(f"⚠️ Erro ao ler {arquivo.name}: {e}")

    if not dados_diarios:
        print(f"❌ Nenhum dado válido pôde ser extraído para {constelacao}.")
        return

    # Criar DataFrame com o resumo diário e ordenar cronologicamente
    df_resumo = pd.DataFrame(dados_diarios)
    df_resumo = df_resumo.sort_values('Data').reset_index(drop=True)

    # Calcular as variações em milímetros
    df_resumo, m_lat, m_lon, m_alt = calcular_variacao_milimetros(df_resumo)

    print(f"🎯 Média {constelacao}: Lat {m_lat:.8f}°, Lon {m_lon:.8f}°, Alt {m_alt:.3f}m")

    # Plotar a Série Temporal
    plt.style.use('ggplot')
    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(12, 9), sharex=True)
    fig.suptitle(f'Série Temporal de Posição PPP - {constelacao}\nVariação Diária (Norte, Leste, Altitude)', fontsize=16, fontweight='bold')

    estilo = {'marker': 'o', 'markersize': 5, 'linewidth': 1.5, 'alpha': 0.8}

    # Norte
    ax1.plot(df_resumo['Data'], df_resumo['dN (mm)'], color='tab:blue', **estilo)
    ax1.set_ylabel('Norte (mm)', fontweight='bold')
    ax1.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)

    # Leste
    ax2.plot(df_resumo['Data'], df_resumo['dE (mm)'], color='tab:orange', **estilo)
    ax2.set_ylabel('Leste (mm)', fontweight='bold')
    ax2.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)

    # Altitude (Up)
    ax3.plot(df_resumo['Data'], df_resumo['dU (mm)'], color='tab:green', **estilo)
    ax3.set_ylabel('Altitude (mm)', fontweight='bold')
    ax3.set_xlabel('Data da Observação', fontweight='bold')
    ax3.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)

    # Formatação do eixo X
    ax3.xaxis.set_major_formatter(mdates.DateFormatter('%d/%b/%Y'))
    plt.gcf().autofmt_xdate()
    
    # Salvar
    caminho_grafico = pasta_resultados / f"Serie_Temporal_{constelacao}.png"
    plt.tight_layout()
    plt.savefig(caminho_grafico, dpi=300, bbox_inches='tight')
    print(f"✅ Gráfico salvo: {caminho_grafico.name}")
    plt.close() # Fecha a figura para não sobrepor com a próxima constelação

def main():
    print("📈 GERADOR DE SÉRIE TEMPORAL DE COORDENADAS (3 Constelações)")

    caminho_salvo = utils.carregar_estado("pasta_rinex_pronta")
    if not caminho_salvo:
        print("❌ Memória não encontrada. Defina a pasta manualmente.")
        return
        
    path_rinex_obs = Path(caminho_salvo)
    if path_rinex_obs.name in ["GPS", "GLONASS", "GPS_GLONASS"]:
        path_rinex_obs = path_rinex_obs.parent

    pasta_resultados = path_rinex_obs.parent / "RESULTADOS_PPP"
    
    arquivos_pos = list(pasta_resultados.glob("*.pos"))
    arquivos_pos = [f for f in arquivos_pos if "events" not in f.name]

    if not arquivos_pos:
        print(f"❌ Nenhum arquivo .pos encontrado na pasta: {pasta_resultados}")
        return

    # Separar os arquivos pelos prefixos no nome
    arq_gps = [f for f in arquivos_pos if f.name.startswith("GPS_") and "GLONASS" not in f.name]
    arq_glo = [f for f in arquivos_pos if f.name.startswith("GLONASS_")]
    arq_gps_glo = [f for f in arquivos_pos if f.name.startswith("GPS_GLONASS_")]

    # Processar e gerar os gráficos um por um
    if arq_gps: 
        processar_e_plotar(arq_gps, "GPS", pasta_resultados)
    if arq_glo: 
        processar_e_plotar(arq_glo, "GLONASS", pasta_resultados)
    if arq_gps_glo: 
        processar_e_plotar(arq_gps_glo, "GPS_GLONASS", pasta_resultados)

    print("\n🎉 Todas as séries temporais foram geradas com sucesso!")

if __name__ == "__main__":
    main()