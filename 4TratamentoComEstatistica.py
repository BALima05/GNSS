import pandas as pd
import numpy as np
import re
import os
import warnings
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

warnings.filterwarnings('ignore')

# ========================================================
# DADOS OFICIAIS DO IBGE POR CONSTELAÇÃO
# ========================================================
DADOS_IBGE = {
    'GPS': {
        2022: {'lat': -23.55564528, 'lon': -46.73031250},
        2023: {'lat': -23.55564525, 'lon': -46.73031267},
        2024: {'lat': -23.55564514, 'lon': -46.73031267}
    },
    'GLONASS': {
        2022: {'lat': -23.55564528, 'lon': -46.73031247},
        2023: {'lat': -23.55564525, 'lon': -46.73031272},
        2024: {'lat': -23.55564514, 'lon': -46.73031272}
    },
    'GPS E GLONASS': {
        2022: {'lat': -23.55564528, 'lon': -46.73031250},
        2023: {'lat': -23.55564525, 'lon': -46.73031269},
        2024: {'lat': -23.55564514, 'lon': -46.73031272}
    }
}

# ========================================================
# FUNÇÕES AUXILIARES
# ========================================================
def dms_to_decimal(dms_str):
    """Converte coordenadas DMS para graus decimais com alta precisão"""
    try:
        parts = re.findall(r"[-+]?\d*\.\d+|\d+", str(dms_str).replace(',', '.'))
        deg = float(parts[0])
        min = float(parts[1]) if len(parts) > 1 else 0.0
        sec = float(parts[2]) if len(parts) > 2 else 0.0
        sign = -1 if any(s in str(dms_str) for s in ['-', 'S', 'W']) else 1
        return round(sign * (abs(deg) + min/60 + sec/3600), 10)
    except:
        return np.nan

def graus_para_metros(d_lat, d_lon, lat):
    """Converte diferença angular para metros"""
    R = 6371000  # Raio da Terra
    d_lat_rad = np.radians(d_lat)
    d_lon_rad = np.radians(d_lon)
    lat_rad = np.radians(lat)
    dx = R * d_lon_rad * np.cos(lat_rad)
    dy = R * d_lat_rad
    return dx, dy

def selecionar_arquivo():
    """Abre o diálogo para seleção do arquivo Excel"""
    root = tk.Tk()
    root.withdraw()
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel com dados GNSS",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    root.destroy()
    return arquivo

# ========================================================
# FUNÇÕES PRINCIPAIS (COM ESTATÍSTICAS)
# ========================================================
def processar_constelacao(df, constelacao):
    """Processa dados de uma constelação específica"""
    try:
        # Verificar colunas obrigatórias
        cols_obrigatorias = ['Data', 'Latitude(gms)', 'Longitude(gms)']
        for col in cols_obrigatorias:
            if col not in df.columns:
                raise ValueError(f"Coluna '{col}' não encontrada")

        # Converter datas para string no formato DD/MM/AAAA
        def formatar_data(data):
            if isinstance(data, str):
                return data
            elif isinstance(data, pd.Timestamp):
                return data.strftime('%d/%m/%Y')
            elif isinstance(data, datetime):
                return data.strftime('%d/%m/%Y')
            else:
                try:
                    return datetime.strptime(str(data), '%d/%m/%Y').strftime('%d/%m/%Y')
                except:
                    return None

        df['Data'] = df['Data'].apply(formatar_data)
        df = df.dropna(subset=['Data'])

        # Converter coordenadas com 10 casas decimais
        df['Latitude Decimal'] = df['Latitude(gms)'].apply(dms_to_decimal)
        df['Longitude Decimal'] = df['Longitude(gms)'].apply(dms_to_decimal)
        
        # Gerar referências com 10 casas decimais
        deslocamentos = calcular_deslocamento_por_constelacao(constelacao)
        if deslocamentos is not None:
            df['Ano'] = df['Data'].apply(lambda x: int(x.split('/')[-1]))
            
            for ano in df['Ano'].unique():
                if ano in DADOS_IBGE.get(constelacao, {}):
                    lat_inicial = DADOS_IBGE[constelacao][ano]['lat']
                    lon_inicial = DADOS_IBGE[constelacao][ano]['lon']
                    delta_lat = deslocamentos.get(ano, {}).get('delta_lat', 0)
                    delta_lon = deslocamentos.get(ano, {}).get('delta_lon', 0)
                    n_dias = 366 if (ano % 4 == 0) else 365
                    
                    mascara = (df['Ano'] == ano)
                    dias_ano = df.loc[mascara, 'Data'].apply(
                        lambda x: datetime.strptime(x, '%d/%m/%Y').timetuple().tm_yday
                    )
                    df.loc[mascara, 'Latitude Referência'] = round(lat_inicial + (dias_ano-1)*(delta_lat/n_dias), 10)
                    df.loc[mascara, 'Longitude Referência'] = round(lon_inicial + (dias_ano-1)*(delta_lon/n_dias), 10)

        # Calcular diferenças com 10 casas decimais
        df['Dif. Latitude (graus)'] = round(df['Latitude Decimal'] - df['Latitude Referência'], 10)
        df['Dif. Longitude (graus)'] = round(df['Longitude Decimal'] - df['Longitude Referência'], 10)
        
        # Converter para metros
        df['Dif. Latitude (m)'] = np.nan
        df['Dif. Longitude (m)'] = np.nan
        for idx, row in df.iterrows():
            if pd.notna(row['Dif. Latitude (graus)']):
                dx, dy = graus_para_metros(
                    row['Dif. Latitude (graus)'],
                    row['Dif. Longitude (graus)'],
                    row['Latitude Decimal']
                )
                df.at[idx, 'Dif. Latitude (m)'] = dy
                df.at[idx, 'Dif. Longitude (m)'] = dx
                
        return df
        
    except Exception as e:
        print(f"ERRO no processamento de {constelacao}: {str(e)}")
        return None

def calcular_deslocamento_por_constelacao(constelacao):
    """Calcula deslocamento anual para uma constelação"""
    dados = DADOS_IBGE.get(constelacao, {})
    if not dados:
        return None
        
    anos = sorted(dados.keys())
    if len(anos) < 2:
        return None
        
    deslocamentos = {}
    for i in range(len(anos)-1):
        ano_atual = anos[i]
        ano_prox = anos[i+1]
        delta_lat = dados[ano_prox]['lat'] - dados[ano_atual]['lat']
        delta_lon = dados[ano_prox]['lon'] - dados[ano_atual]['lon']
        deslocamentos[ano_atual] = {'delta_lat': delta_lat, 'delta_lon': delta_lon}
        
    return deslocamentos

def gerar_estatisticas_consolidadas(resultados):
    """Gera estatísticas seguindo exatamente o método do código original"""
    metricas = [
        'Latitude(gms) em decimal',
        'Latitude Referência',
        'Dif. Latitude (m)',
        'Sigma Latitude (95%) (m)',
        'Longitude(gms) em decimal',
        'Longitude Referência',
        'Dif. Longitude (m)',
        'Sigma Longitude (95%) (m)',
        'Altitude Geométrica (m)',
        'Sigma Alt. Geo. (95%) (m)',
        'Resíduos da pseudo distância GPS (m)',
        'Resíduos da pseudo distância GLONASS (m)',
        'Resíduos da fase da portadora GPS (cm)',
        'Resíduos da fase da portadora GLONASS (cm)'
    ]
    
    estatisticas = []
    
    for metrica in metricas:
        linha_metrica = {'Métrica': metrica}
        
        for constelacao in ['GPS', 'GLONASS', 'GPS E GLONASS']:
            df = resultados.get(constelacao)
            
            if df is None:
                linha_metrica[f'{constelacao} (Média)'] = '-'
                linha_metrica[f'{constelacao} (Desvio Padrão)'] = '-'
                continue
                
            def to_float(value):
                try:
                    return float(str(value).replace(',', '.'))
                except:
                    return np.nan
            
            if 'Sigma' in metrica:
                col_name = metrica
            else:
                metric_map = {
                    'Latitude(gms) em decimal': 'Latitude Decimal',
                    'Latitude Referência': 'Latitude Referência',
                    'Dif. Latitude (m)': 'Dif. Latitude (m)',
                    'Longitude(gms) em decimal': 'Longitude Decimal',
                    'Longitude Referência': 'Longitude Referência',
                    'Dif. Longitude (m)': 'Dif. Longitude (m)',
                    'Altitude Geométrica (m)': 'Altitude Geométrica (m)',
                    'Resíduos da pseudo distância GPS (m)': 'Resíduos da pseudo distância GPS (m)',
                    'Resíduos da pseudo distância GLONASS (m)': 'Resíduos da pseudo distância GLONASS (m)',
                    'Resíduos da fase da portadora GPS (cm)': 'Resíduos da fase da portadora GPS (cm)',
                    'Resíduos da fase da portadora GLONASS (cm)': 'Resíduos da fase da portadora GLONASS (cm)'
                }
                col_name = metric_map.get(metrica, None)
            
            if col_name and col_name in df.columns:
                valores = df[col_name].apply(to_float).dropna()
                if len(valores) > 0:
                    linha_metrica[f'{constelacao} (Média)'] = round(valores.mean(), 10)
                    linha_metrica[f'{constelacao} (Desvio Padrão)'] = round(valores.std(), 10)
                else:
                    linha_metrica[f'{constelacao} (Média)'] = '-'
                    linha_metrica[f'{constelacao} (Desvio Padrão)'] = '-'
            else:
                linha_metrica[f'{constelacao} (Média)'] = '-'
                linha_metrica[f'{constelacao} (Desvio Padrão)'] = '-'
        
        estatisticas.append(linha_metrica)
    
    df_estat = pd.DataFrame(estatisticas)
    
    colunas_ordenadas = ['Métrica']
    for constelacao in ['GPS', 'GLONASS', 'GPS E GLONASS']:
        colunas_ordenadas.extend([f'{constelacao} (Média)', f'{constelacao} (Desvio Padrão)'])
    
    return df_estat[colunas_ordenadas]

def formatar_estatisticas(df_estat):
    """Formata o DataFrame de estatísticas para melhor visualização"""
    for col in df_estat.columns:
        if col != 'Métrica':
            df_estat[col] = df_estat[col].apply(
                lambda x: f"{x:.10f}" if isinstance(x, (int, float)) and not np.isnan(x) else 'N/D'
            )
    
    return df_estat

# ========================================================
# PROCESSAMENTO PRINCIPAL
# ========================================================
def processar_planilha_principal(input_file):
    """Processamento completo do arquivo de entrada"""
    CONSTELACOES = ['GPS', 'GLONASS', 'GPS E GLONASS']
    resultados = {}
    
    try:
        with pd.ExcelFile(input_file) as xls:
            for constelacao in CONSTELACOES:
                if constelacao not in xls.sheet_names:
                    print(f"AVISO: Aba '{constelacao}' não encontrada")
                    resultados[constelacao] = None
                    continue
                    
                df = pd.read_excel(xls, sheet_name=constelacao)
                df_processado = processar_constelacao(df, constelacao)
                
                if df_processado is not None:
                    cols_adicionais = [c for c in df.columns if c not in df_processado.columns]
                    for col in cols_adicionais:
                        df_processado[col] = df[col]
                
                resultados[constelacao] = df_processado
    
    except Exception as e:
        print(f"ERRO CRÍTICO: {str(e)}")
        return None, None
        
    df_estatisticas = gerar_estatisticas_consolidadas(resultados)
    df_estatisticas_formatado = formatar_estatisticas(df_estatisticas)
    
    return resultados, df_estatisticas_formatado

# ========================================================
# EXECUÇÃO PRINCIPAL
# ========================================================
if __name__ == "__main__":
    print("=== Sistema de Análise GNSS ===")
    print("Processamento de dados GPS, GLONASS e combinado\n")
    
    input_file = selecionar_arquivo()
    if not input_file:
        print("Processamento cancelado.")
    else:
        print(f"\nProcessando arquivo: {os.path.basename(input_file)}")
        
        resultados, estatisticas = processar_planilha_principal(input_file)
        
        if resultados:
            output_file = os.path.join(
                os.path.dirname(input_file),
                f"Resultados_GNSS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            
            try:
                with pd.ExcelWriter(output_file) as writer:
                    for constelacao, df in resultados.items():
                        if df is not None:
                            df.to_excel(
                                writer,
                                sheet_name=constelacao[:31],
                                index=False,
                                float_format="%.10f"  # Alterado para 10 casas decimais
                            )
                    
                    estatisticas.to_excel(
                        writer,
                        sheet_name='ESTATISTICAS',
                        index=False
                    )
                    
                    for sheet in writer.sheets:
                        worksheet = writer.sheets[sheet]
                        for col in worksheet.columns:
                            max_length = max(
                                len(str(cell.value)) for cell in col
                            )
                            worksheet.column_dimensions[col[0].column_letter].width = max_length + 2
                
                print(f"\nProcessamento concluído com sucesso!")
                print(f"Resultados salvos em: {output_file}")
                
                if os.name == 'nt':
                    try:
                        os.startfile(output_file)
                    except:
                        print("Não foi possível abrir o arquivo automaticamente.")
            
            except Exception as e:
                print(f"\nERRO ao salvar resultados: {str(e)}")