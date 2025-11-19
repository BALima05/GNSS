import pandas as pd
import numpy as np
import re
import os
import warnings
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

warnings.filterwarnings('ignore')

# Fatores de correção aumentados para efeitos de demonstração
FATORES = {
    2022: {
        'GPS': {'lat': 8.38e-11, 'lon': -4.57e-10},  # Aumentado em 100x
        'GLONASS': {'lat': 8.38e-11, 'lon': -6.85e-10},
        'GPS E GLONASS': {'lat': 8.38e-11, 'lon': -5.3e-10}
    },
    2023: {
        'GPS': {'lat': 2.97e-10, 'lon': 0.0},
        'GLONASS': {'lat': 2.97e-10, 'lon': 0.0},
        'GPS E GLONASS': {'lat': 2.97e-10, 'lon': -7.62E-11}
    }
}

def dms_to_decimal(dms_str):
    """Converte coordenadas DMS para graus decimais com alta precisão"""
    try:
        if pd.isna(dms_str):
            return np.nan
            
        # Extrair componentes numéricas
        parts = re.findall(r"[-+]?\d*\.\d+|\d+", str(dms_str).replace(',', '.'))
        if not parts:
            return np.nan
            
        # Converter para float
        deg = float(parts[0])
        min = float(parts[1]) if len(parts) > 1 else 0.0
        sec = float(parts[2]) if len(parts) > 2 else 0.0
        
        # Determinar sinal
        sign = -1 if any(s in str(dms_str) for s in ['-', 'S', 'W']) else 1
        
        return sign * (abs(deg) + min/60 + sec/3600)
    except Exception as e:
        print(f"Erro na conversão: {dms_str} - {e}")
        return np.nan

def graus_para_metros(d_lat, d_lon, lat):
    """Converte diferença em graus para metros"""
    # Raio da Terra em metros
    R = 6371000
    
    # Conversão para radianos
    d_lat_rad = np.radians(d_lat)
    d_lon_rad = np.radians(d_lon)
    lat_rad = np.radians(lat)
    
    # Cálculo das distâncias
    dx = R * d_lon_rad * np.cos(lat_rad)
    dy = R * d_lat_rad
    
    return dx, dy

def selecionar_arquivo(titulo="Selecionar arquivo", tipo_arquivo=[("Excel files", "*.xlsx")]):
    """Abre uma janela para seleção de arquivo"""
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    arquivo = filedialog.askopenfilename(title=titulo, filetypes=tipo_arquivo)
    root.destroy()
    return arquivo

def processar_planilha():
    # Selecionar arquivo de entrada
    input_file = selecionar_arquivo("Selecione o arquivo Excel de entrada")
    if not input_file:
        print("Nenhum arquivo selecionado. Operação cancelada.")
        return
    
    # Gerar nome do arquivo de saída no mesmo diretório
    dir_path = os.path.dirname(input_file)
    file_name = os.path.basename(input_file)
    name, ext = os.path.splitext(file_name)
    output_file = os.path.join(dir_path, f"{name}_processado{ext}")
    
    # Abas a serem processadas
    CONSTELACOES = ['GPS', 'GLONASS', 'GPS E GLONASS']
    
    # Carregar dados
    dados_originais = {}
    with pd.ExcelFile(input_file) as xls:
        for constelacao in CONSTELACOES:
            if constelacao in xls.sheet_names:
                # Ler dados mantendo datas como strings
                df = pd.read_excel(xls, sheet_name=constelacao)
                
                # Tratamento direto de datas
                datas = []
                for data in df['Data']:
                    if isinstance(data, str) and '/' in data:
                        # Já está no formato DD/MM/AAAA
                        datas.append(data)
                    elif isinstance(data, datetime):
                        # Converter datetime para string DD/MM/AAAA
                        datas.append(data.strftime('%d/%m/%Y'))
                    else:
                        # Tentar converter números ou outros formatos
                        try:
                            if isinstance(data, float) or isinstance(data, int):
                                # Converter número de série do Excel para data
                                dt = pd.Timestamp('1899-12-30') + pd.Timedelta(days=data)
                                datas.append(dt.strftime('%d/%m/%Y'))
                            else:
                                datas.append(str(data))
                        except:
                            datas.append('DATA INVÁLIDA')
                
                df['Data'] = datas
                dados_originais[constelacao] = df
    
    # =================================================================
    # Impressão das datas
    # =================================================================
    print("\nDatas dos arquivos na planilha:")
    for constelacao, df in dados_originais.items():
        # Remover duplicatas e manter ordem original
        datas_unicas = []
        for data in df['Data']:
            if data not in datas_unicas and 'DATA INVÁLIDA' not in data:
                datas_unicas.append(data)
        
        if datas_unicas:
            print(f"\n{constelacao} ({len(datas_unicas)} datas):")
            print(", ".join(datas_unicas))
        else:
            print(f"\n{constelacao}: Nenhuma data válida encontrada")
    
    # Criar ExcelWriter
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Dicionário para armazenar dados tratados
        dados_tratados = {}
        
        # ETAPA 1: Tratamento de coordenadas
        for constelacao in CONSTELACOES:
            if constelacao not in dados_originais:
                continue
                
            df = dados_originais[constelacao].copy()
            
            # Criar dataframe para tratamento
            df_trat = df[['Nome do arquivo', 'Tipo', 'Data']].copy()
            
            # Adicionar colunas originais como strings
            df_trat['Latitude(gms)'] = df['Latitude(gms)'].astype(str)
            df_trat['Longitude(gms)'] = df['Longitude(gms)'].astype(str)
            
            # Inicializar colunas calculadas
            df_trat['Latitude Decimal'] = np.nan
            df_trat['Longitude Decimal'] = np.nan
            df_trat['Latitude Ajustada'] = np.nan
            df_trat['Dif. Latitude (graus)'] = np.nan
            df_trat['Longitude Ajustada'] = np.nan
            df_trat['Dif. Longitude (graus)'] = np.nan
            df_trat['Dif. Latitude (m)'] = np.nan
            df_trat['Dif. Longitude (m)'] = np.nan
            
            # Extrair ano da data para aplicar correções
            anos = []
            for data_str in df['Data']:
                try:
                    # Extrair ano diretamente da string DD/MM/AAAA
                    ano = int(data_str.split('/')[-1])
                    anos.append(ano)
                except:
                    anos.append(2022)  # Ano padrão se falhar
            
            df['Ano'] = anos
            
            # Processar cada linha
            for idx, row in df.iterrows():
                ano = df.at[idx, 'Ano']
                constelacao_atual = constelacao
                
                # Obter fatores de correção ANUAL
                fator = FATORES.get(ano, {}).get(constelacao_atual, {'lat': 0, 'lon': 0})
                delta_lat_anual = fator['lat']  # Variação total anual em graus
                delta_lon_anual = fator['lon']
                
                # Converter data para dia do ano (1 a 365/366)
                data_str = row['Data']
                try:
                    dia_ano = datetime.strptime(data_str, '%d/%m/%Y').timetuple().tm_yday
                except:
                    dia_ano = 1  # Padrão se falhar
                
                # Calcular variação DIÁRIA (Δφ e Δλ)
                n_dias = 366 if ano % 4 == 0 else 365  # Bissexto
                delta_lat_diario = delta_lat_anual / n_dias
                delta_lon_diario = delta_lon_anual / n_dias
                
                # Converter coordenadas originais
                lat_orig = dms_to_decimal(row['Latitude(gms)'])
                lon_orig = dms_to_decimal(row['Longitude(gms)'])
                
                # Calcular coordenada estimada para o dia atual (modelo linear)
                lat_estimada = lat_orig + (dia_ano - 1) * delta_lat_diario
                lon_estimada = lon_orig + (dia_ano - 1) * delta_lon_diario
                
                # Calcular diferenças em graus (observado - estimado)
                dif_lat = lat_orig - lat_estimada
                dif_lon = lon_orig - lon_estimada
                
                # Calcular diferenças em metros
                dx, dy = graus_para_metros(dif_lat, dif_lon, lat_orig)
                
                # Armazenar resultados
                df_trat.at[idx, 'Latitude Decimal'] = lat_orig
                df_trat.at[idx, 'Longitude Decimal'] = lon_orig
                df_trat.at[idx, 'Latitude Ajustada'] = lat_estimada
                df_trat.at[idx, 'Longitude Ajustada'] = lon_estimada
                df_trat.at[idx, 'Dif. Latitude (graus)'] = dif_lat
                df_trat.at[idx, 'Dif. Longitude (graus)'] = dif_lon
                df_trat.at[idx, 'Dif. Latitude (m)'] = dy
                df_trat.at[idx, 'Dif. Longitude (m)'] = dx
            
            # Salvar aba de tratamento
            df_trat.to_excel(writer, sheet_name=f"Tratamento {constelacao}", index=False)
            dados_tratados[constelacao] = df_trat
        
        # Restante do código (ETAPA 2: Estatísticas) permanece igual...
        # ETAPA 2: Estatísticas consolidadas
        # Criar estrutura da tabela de estatísticas
        metricas = [
            'Latitude(gms) em decimal',
            'Latitude Ajustada em decimal',
            'Dif. Latitude (m)',  # Mantido
            'Sigma Latitude (95%) (m)',
            'Longitude(gms) em decimal',
            'Longitude Ajustada em decimal',
            'Dif. Longitude (m)',  # Mantido
            'Sigma Longitude (95%) (m)',
            'Altitude Geométrica (m)',
            'Sigma Alt. Geo. (95%) (m)',
            'Resíduos da pseudo distância GPS (m)',
            'Resíduos da pseudo distância GLONASS (m)',
            'Resíduos da fase da portadora GPS (cm)',
            'Resíduos da fase da portadora GLONASS (cm)'
        ]
        
        # Criar dataframe de estatísticas
        stats_df = pd.DataFrame({
            'Métrica': metricas,
            'GPS Média': np.nan,
            'GPS Desvio Padrão': np.nan,
            'GLONASS Média': np.nan,
            'GLONASS Desvio Padrão': np.nan,
            'GPS E GLONASS Média': np.nan,
            'GPS E GLONASS Desvio Padrão': np.nan
        })
        
        # Função para converter valores numéricos
        def to_float(value):
            try:
                return float(str(value).replace(',', '.'))
            except:
                return np.nan
        
        # Calcular estatísticas para cada constelação
        for constelacao in CONSTELACOES:
            if constelacao not in dados_originais or constelacao not in dados_tratados:
                continue
                
            df_orig = dados_originais[constelacao]
            df_trat = dados_tratados[constelacao]
            
            # Mapeamento de métricas para colunas/fontes
            metric_map = {
                'Latitude(gms) em decimal': 
                    df_trat['Latitude Decimal'],
                'Latitude Ajustada em decimal': 
                    df_trat['Latitude Ajustada'],
                'Dif. Latitude (m)': 
                    df_trat['Dif. Latitude (m)'],
                'Sigma Latitude (95%) (m)': 
                    df_orig['Sigma Latitude (95%) (m)'].apply(to_float),
                'Longitude(gms) em decimal': 
                    df_trat['Longitude Decimal'],
                'Longitude Ajustada em decimal': 
                    df_trat['Longitude Ajustada'],
                'Dif. Longitude (m)': 
                    df_trat['Dif. Longitude (m)'],
                'Sigma Longitude (95%) (m)': 
                    df_orig['Sigma Longitude (95%) (m)'].apply(to_float),
                'Altitude Geométrica (m)': 
                    df_orig['Altitude Geométrica (m)'].apply(to_float),
                'Sigma Alt. Geo. (95%) (m)': 
                    df_orig['Sigma Alt. Geo. (95%) (m)'].apply(to_float),
                'Resíduos da pseudo distância GPS (m)': 
                    df_orig['Resíduos da pseudo distância GPS (m)'].apply(to_float) 
                    if 'Resíduos da pseudo distância GPS (m)' in df_orig.columns 
                    else pd.Series([np.nan] * len(df_orig)),
                'Resíduos da pseudo distância GLONASS (m)': 
                    df_orig['Resíduos da pseudo distância GLONASS (m)'].apply(to_float) 
                    if 'Resíduos da pseudo distância GLONASS (m)' in df_orig.columns 
                    else pd.Series([np.nan] * len(df_orig)),
                'Resíduos da fase da portadora GPS (cm)': 
                    df_orig['Resíduos da fase da portadora GPS (cm)'].apply(to_float) 
                    if 'Resíduos da fase da portadora GPS (cm)' in df_orig.columns 
                    else pd.Series([np.nan] * len(df_orig)),
                'Resíduos da fase da portadora GLONASS (cm)': 
                    df_orig['Resíduos da fase da portadora GLONASS (cm)'].apply(to_float) 
                    if 'Resíduos da fase da portadora GLONASS (cm)' in df_orig.columns 
                    else pd.Series([np.nan] * len(df_orig))
            }
            
            # Calcular estatísticas
            for idx, row in stats_df.iterrows():
                metrica = row['Métrica']
                valores = metric_map.get(metrica, None)
                
                if valores is None:
                    continue
                    
                # Calcular média e desvio padrão
                try:
                    media = valores.mean()
                    desvio = valores.std()
                except Exception as e:
                    print(f"Erro ao calcular estatísticas para {metrica} ({constelacao}): {e}")
                    media = np.nan
                    desvio = np.nan
                
                # Atualizar dataframe
                col_prefix = constelacao
                stats_df.loc[idx, f'{col_prefix} Média'] = media
                stats_df.loc[idx, f'{col_prefix} Desvio Padrão'] = desvio
        
        # Substituir NaN por '-' onde não há dados
        stats_df = stats_df.fillna('-')
        
        # Salvar aba de estatísticas
        stats_df.to_excel(writer, sheet_name='Estatísticas', index=False)
    
    print(f"\nProcessamento concluído! Resultados salvos em: {output_file}")
    # Abrir o arquivo gerado automaticamente
    if os.name == 'nt':  # Somente para Windows
        os.startfile(output_file)

# Executar processamento
if __name__ == "__main__":
    processar_planilha()