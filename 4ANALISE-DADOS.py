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
    Lê os arquivos, filtra os dados por um limite de qualidade absoluto (< 10 cm),
    extrai a melhor coordenada diária e gera o gráfico.
    """
    print(f"\n🔍 Processando {len(arquivos)} arquivos para a constelação: {constelacao}...")
    
    colunas_padrao = ['Date', 'Time', 'Lat', 'Lon', 'Height', 'Q', 'ns', 'sdn', 'sde', 'sdu', 'sdne', 'sdeu', 'sdun', 'age', 'ratio']
    dados_diarios = []

    # --- LIMITE ABSOLUTO DE QUALIDADE (metros) ---
    LIMITE_ERRO_METROS = 0.10

    for arquivo in arquivos:
        try:
            df_temp = pd.read_csv(arquivo, comment='%', sep=r'\s+', names=colunas_padrao, encoding='latin1')
            
            if df_temp.empty:
                continue

            cols_calc = ['Lat', 'Lon', 'Height', 'Q', 'sdn', 'sde', 'sdu', 'ns']
            for col in cols_calc:
                df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce')
            
            df_temp = df_temp.dropna(subset=cols_calc)

            # 1. Ignorar SPP (Q=5)
            df_temp = df_temp[df_temp['Q'].isin([1, 2, 6])]
            
            # 2. Calcular a incerteza 3D
            df_temp['sd_3d'] = (df_temp['sdn']**2 + df_temp['sde']**2 + df_temp['sdu']**2)**0.5

            # 3. FILTRO ABSOLUTO: Excluir tudo que tem erro maior que o limite
            df_temp = df_temp[df_temp['sd_3d'] <= LIMITE_ERRO_METROS]

            if df_temp.empty:
                print(f"      ⚠️ {arquivo.name}: Ignorado (Nenhuma época alcançou precisão 3D < {LIMITE_ERRO_METROS*100:.0f}cm)")
                continue

            # 4. Pegar a nata dos 5% melhores, MAS agora dentro dos que já passaram no teste do limite
            df_temp = df_temp.sort_values('sd_3d')
            n_pontos = max(1, int(len(df_temp) * 0.05))
            df_melhores = df_temp.head(n_pontos)

            data_dia = pd.to_datetime(df_temp['Date'].iloc[0])
            
            dados_diarios.append({
                'Data': data_dia,
                'Lat': df_melhores['Lat'].median(),
                'Lon': df_melhores['Lon'].median(),
                'Alt': df_melhores['Height'].median()
            })
        except Exception as e:
            print(f"⚠️ Erro ao processar {arquivo.name}: {e}")

    if not dados_diarios:
        print(f"❌ Nenhum dado válido pôde ser extraído para {constelacao}.")
        return

    # Criar DataFrame final
    df_resumo = pd.DataFrame(dados_diarios)
    df_resumo = df_resumo.sort_values('Data').reset_index(drop=True)

    # Calcular as variações reais
    df_resumo, m_lat, m_lon, m_alt = calcular_variacao_milimetros(df_resumo)

    # Cálculos Estatísticos
    std_n, std_e, std_u = df_resumo['dN (mm)'].std(), df_resumo['dE (mm)'].std(), df_resumo['dU (mm)'].std()
    rmse_n = (df_resumo['dN (mm)']**2).mean()**0.5
    rmse_e = (df_resumo['dE (mm)']**2).mean()**0.5
    rmse_u = (df_resumo['dU (mm)']**2).mean()**0.5

    print(f"🎯 Média {constelacao}: Lat {m_lat:.8f}°, Lon {m_lon:.8f}°, Alt {m_alt:.3f}m")
    print(f"📊 ESTATÍSTICAS DA SÉRIE (Variação em milímetros):")
    print(f"  NORTE -> Desvio Padrão: {std_n:.2f} | RMSE: {rmse_n:.2f} | Min: {df_resumo['dN (mm)'].min():.2f} | Máx: {df_resumo['dN (mm)'].max():.2f}")
    print(f"  LESTE -> Desvio Padrão: {std_e:.2f} | RMSE: {rmse_e:.2f} | Min: {df_resumo['dE (mm)'].min():.2f} | Máx: {df_resumo['dE (mm)'].max():.2f}")
    print(f"  ALT   -> Desvio Padrão: {std_u:.2f} | RMSE: {rmse_u:.2f} | Min: {df_resumo['dU (mm)'].min():.2f} | Máx: {df_resumo['dU (mm)'].max():.2f}")

    # Salvar tabela Excel/CSV
    pasta_tabelas = pasta_resultados / "Tabelas_Estatisticas"
    os.makedirs(pasta_tabelas, exist_ok=True)
    arquivo_csv = pasta_tabelas / f"Dados_Serie_{constelacao}.csv"
    df_resumo.to_csv(arquivo_csv, index=False, sep=';', decimal=',')

    # Plotar os gráficos
    plt.style.use('ggplot')
    fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(12, 9), sharex=True)
    fig.suptitle(f'Série Temporal de Posição PPP - {constelacao}\nVariação Diária (Norte, Leste, Altitude)', fontsize=16, fontweight='bold')

    estilo = {'marker': 'o', 'markersize': 5, 'linewidth': 1.5, 'alpha': 0.8}

    LIMITE_NORTE = 100
    LIMITE_LESTE = 100
    LIMITE_ALTITUDE = 200

    ax1.plot(df_resumo['Data'], df_resumo['dN (mm)'], color='tab:blue', **estilo)
    ax1.set_ylabel('Norte (mm)', fontweight='bold')
    ax1.set_ylim(-LIMITE_NORTE, LIMITE_NORTE)         ## Define os limites do eixo para a Latitude
    ax1.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)

    # Caixa de Estatísticas no Gráfico
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
    ax2.set_ylim(-LIMITE_LESTE, LIMITE_LESTE)         ## Define os limites do eixo para a Longitude
    ax2.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)

    ax3.plot(df_resumo['Data'], df_resumo['dU (mm)'], color='tab:green', **estilo)
    ax3.set_ylabel('Altitude (mm)', fontweight='bold')
    ax3.set_xlabel('Data da Observação', fontweight='bold')
    ax3.set_ylim(-LIMITE_ALTITUDE, LIMITE_ALTITUDE)        ## Define os limites do eixo para a Altitude
    ax3.axhline(0, color='black', linestyle='-', linewidth=1, alpha=0.5)

    ax3.xaxis.set_major_formatter(mdates.DateFormatter('%d/%b/%Y'))
    plt.gcf().autofmt_xdate()
    
    # Salvar gráfico
    pasta_graficos = pasta_resultados / "Graficos_Serie_Temporal"
    os.makedirs(pasta_graficos, exist_ok=True)
    caminho_grafico = pasta_graficos / f"Serie_Temporal_{constelacao}.png"
    plt.tight_layout()
    plt.savefig(caminho_grafico, dpi=300, bbox_inches='tight')
    print(f"✅ Gráfico salvo: {caminho_grafico.name}")
    plt.close()

def main():
    print("📈 GERADOR DE SÉRIE TEMPORAL COM GUILHOTINA ABSOLUTA (3 Constelações)")

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

    arq_gps = [f for f in arquivos_pos if f.name.startswith("GPS_") and "GLONASS" not in f.name]
    arq_glo = [f for f in arquivos_pos if f.name.startswith("GLONASS_")]
    arq_gps_glo = [f for f in arquivos_pos if f.name.startswith("GPS_GLONASS_")]

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