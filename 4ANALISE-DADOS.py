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

    # Diferença em relação à média convertida para milímetros (* 1000)
    df['dN (mm)'] = (df['Lat'] - media_lat) * fator_lat * 1000
    df['dE (mm)'] = (df['Lon'] - media_lon) * fator_lon * 1000
    df['dU (mm)'] = (df['Alt'] - media_alt) * 1000

    return df, media_lat, media_lon, media_alt

def main():
    print("📈 GERADOR DE SÉRIE TEMPORAL DE COORDENADAS (Inter-Dias)")

    # 1. Localizar os resultados
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

    print(f"🔍 Extraindo coordenadas diárias de {len(arquivos_pos)} arquivos. Aguarde...")

    # 2. Ler todos os arquivos e extrair a Mediana de cada dia
    colunas_padrao = ['Date', 'Time', 'Lat', 'Lon', 'Height', 'Q', 'ns', 'sdn', 'sde', 'sdu', 'sdne', 'sdeu', 'sdun', 'age', 'ratio']
    dados_diarios = []

    for arquivo in arquivos_pos:
        try:
            # Lê o arquivo ignorando os cabeçalhos
            df_temp = pd.read_csv(arquivo, comment='%', delim_whitespace=True, names=colunas_padrao)
            
            if df_temp.empty:
                continue

            # Pega a data do primeiro registro do arquivo
            data_dia = pd.to_datetime(df_temp['Date'].iloc[0])
            
            # Usa a mediana para ignorar a fase de convergência inicial do PPP
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
        print("❌ Nenhum dado válido pôde ser extraído dos arquivos.")
        return

    # 3. Criar DataFrame com o resumo diário e ordenar cronologicamente
    df_resumo = pd.DataFrame(dados_diarios)
    df_resumo = df_resumo.sort_values('Data').reset_index(drop=True)

    # 4. Calcular as variações em milímetros
    df_resumo, m_lat, m_lon, m_alt = calcular_variacao_milimetros(df_resumo)

    print("\n==================================================")
    print("🎯 COORDENADA MÉDIA DA SÉRIE TEMPORAL:")
    print(f"   Latitude:  {m_lat:.8f}°")
    print(f"   Longitude: {m_lon:.8f}°")
    print(f"   Altitude:  {m_alt:.3f} metros")
    print("==================================================\n")

    # 5. Plotar a Série Temporal
    print("🎨 Gerando gráfico da série temporal...")
    plt.style.use('ggplot')
    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(12, 9), sharex=True)
    fig.suptitle('Série Temporal de Posição PPP\nVariação Diária (Norte, Leste, Altitude)', fontsize=16, fontweight='bold')

    # Configuração dos marcadores (bolinhas com linhas conectando)
    estilo = {'marker': 'o', 'markersize': 5, 'linewidth': 1.5, 'alpha': 0.8}

    # Norte
    ax1.plot(df_resumo['Data'], df_resumo['dN (mm)'], color='tab:blue', **estilo)
    ax1.set_ylabel('Norte (mm)', fontweight='bold')
    ax1.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)
    ax1.grid(True, linestyle=':', alpha=0.7)

    # Leste
    ax2.plot(df_resumo['Data'], df_resumo['dE (mm)'], color='tab:orange', **estilo)
    ax2.set_ylabel('Leste (mm)', fontweight='bold')
    ax2.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)
    ax2.grid(True, linestyle=':', alpha=0.7)

    # Altitude (Up)
    ax3.plot(df_resumo['Data'], df_resumo['dU (mm)'], color='tab:green', **estilo)
    ax3.set_ylabel('Altitude (mm)', fontweight='bold')
    ax3.set_xlabel('Data da Observação', fontweight='bold')
    ax3.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)
    ax3.grid(True, linestyle=':', alpha=0.7)

    # Formatação do eixo X para mostrar datas bonitas
    ax3.xaxis.set_major_formatter(mdates.DateFormatter('%d/%b/%Y'))
    plt.gcf().autofmt_xdate() # Rotaciona as datas para não sobrepor
    
    # Salvar e mostrar
    caminho_grafico = pasta_resultados / "Serie_Temporal_PPP.png"
    plt.tight_layout()
    plt.savefig(caminho_grafico, dpi=300, bbox_inches='tight')
    print(f"✅ Gráfico salvo em alta resolução: {caminho_grafico}")
    
    plt.show()

if __name__ == "__main__":
    main()