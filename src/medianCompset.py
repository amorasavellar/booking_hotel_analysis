import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from datetime import datetime, timedelta
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, Border, Side



# Configuração da página Streamlit
st.set_page_config(page_title="price comparison", layout="wide")

def clean_price(price):
    if pd.isna(price) or price == '':
        return None
    if isinstance(price, str):
        # Remover caracteres não numéricos e converter para float
        cleaned = ''.join(filter(str.isdigit, price))
        return float(cleaned) if cleaned else None
    return float(price)

# Diretório onde estão os arquivos Excel
directory = "C:/Users/ribei/Documents/RegiOtels/Dashboard-estatistica/DetailedPrices"

# Listas para armazenar os dataframes
competitors_dfs = []
khaolak_dfs = []

print(f"Buscando arquivos em: {directory}")

# Ler todos os arquivos Excel no diretório
for filename in os.listdir(directory):
    if filename.endswith(".xlsx") and "detailed_prices" in filename:
        file_path = os.path.join(directory, filename)
        print(f"/nProcessando arquivo: {filename}")
        try:
            df = pd.read_excel(file_path)
            df['Hotel'] = filename.split('_')[0]  # Extrair nome do hotel do arquivo
            
            # Mostrar as primeiras linhas e informações sobre a coluna 'Price'
            print(f"Primeiras 5 linhas do DataFrame:\n{df.head()}")
            print(f"Informações sobre a coluna 'Price':\n{df['Price'].describe()}")
            print(f"Valores únicos na coluna 'Price': {df['Price'].unique()}")
            
            # Limpar e converter a coluna 'Price'
            df['Price'] = df['Price'].apply(clean_price)
            
            # Remover linhas com preços nulos
            df_cleaned = df.dropna(subset=['Price'])
            print(f"Linhas antes da limpeza: {len(df)}, após limpeza: {len(df_cleaned)}")
            
            if "khaolak" in filename.lower():
                khaolak_dfs.append(df_cleaned)
                print(f"Adicionado à lista khaolak: {filename}")
            else:
                competitors_dfs.append(df_cleaned)
                print(f"Adicionado à lista de competidores: {filename}")
        except Exception as e:
            print(f"Erro ao processar {filename}: {str(e)}")

print(f"\nTotal de arquivos de competidores: {len(competitors_dfs)}")
print(f"Total de arquivos de Khaolak: {len(khaolak_dfs)}")

# Verificar se há dados para processar
if not competitors_dfs and not khaolak_dfs:
    print("Nenhum arquivo válido encontrado. Verifique o diretório e os nomes dos arquivos.")
    exit()

# Combinar os dataframes
if competitors_dfs:
    competitors_df = pd.concat(competitors_dfs, ignore_index=True)
    print(f"\nShape do DataFrame de competidores: {competitors_df.shape}")
    print(f"Tipos de dados: {competitors_df.dtypes}")
    print(f"Resumo estatístico dos preços dos competidores:\n{competitors_df['Price'].describe()}")
else:
    print("Nenhum dado de competidores para processar.")

if khaolak_dfs:
    khaolak_df = pd.concat(khaolak_dfs, ignore_index=True)
    print(f"\nShape do DataFrame de Khaolak: {khaolak_df.shape}")
    print(f"Tipos de dados: {khaolak_df.dtypes}")
    print(f"Resumo estatístico dos preços de Khaolak:\n{khaolak_df['Price'].describe()}")
else:
    print("Nenhum dado de Khaolak para processar.")

def calculate_stats(df, start_date, end_date):
    mask = (df['Date'] >= start_date) & (df['Date'] <= end_date)
    period_data = df.loc[mask]
    return {
        'mean': period_data['Price'].mean(),
        'min': period_data['Price'].min(),
        'max': period_data['Price'].max(),
        'median': period_data['Price'].median()
    }


st.title("price comparison: Khaolak vs competitors")



# Slider para selecionar o período
periods = {
    "1 Mês": 30,
    "3 Meses": 90,
    "6 Meses": 180,
    "1 Ano": 365
}
selected_period = st.selectbox("Selecione o período de visualização:", list(periods.keys()))

# Converter a coluna 'Date' para datetime se ainda não estiver
khaolak_df['Date'] = pd.to_datetime(khaolak_df['Date'])
competitors_df['Date'] = pd.to_datetime(competitors_df['Date'])

# Definir end_date como a data máxima dos dataframes
end_date = max(khaolak_df['Date'].max(), competitors_df['Date'].max())

# Calcular start_date baseado no período selecionado
start_date = end_date - timedelta(days=periods[selected_period])

# Filtrar dados para o período selecionado
khaolak_filtered = khaolak_df[(khaolak_df['Date'] >= start_date) & (khaolak_df['Date'] <= end_date)]
competitors_filtered = competitors_df[(competitors_df['Date'] >= start_date) & (competitors_df['Date'] <= end_date)]
# Filtrar dados para o período selecionado
khaolak_filtered = khaolak_df[(khaolak_df['Date'] >= start_date) & (khaolak_df['Date'] <= end_date)]
competitors_filtered = competitors_df[(competitors_df['Date'] >= start_date) & (competitors_df['Date'] <= end_date)]

# Calcular medianas
khaolak_median = khaolak_filtered.groupby('Date')['Price'].median().reset_index()
competitors_median = competitors_filtered.groupby('Date')['Price'].median().reset_index()

# Calcular estatísticas para o período
khaolak_stats = calculate_stats(khaolak_filtered, start_date, end_date)
competitors_stats = calculate_stats(competitors_filtered, start_date, end_date)

# Calcular diferença percentual
diff_percentage = ((khaolak_stats['median'] - competitors_stats['median']) / competitors_stats['median']) * 100

# Criar o gráfico
fig = make_subplots(specs=[[{"secondary_y": True}]])


# Converter a coluna 'Date' para datetime se ainda não estiver
khaolak_df['Date'] = pd.to_datetime(khaolak_df['Date'])
competitors_df['Date'] = pd.to_datetime(competitors_df['Date'])

# Encontrar a data mínima e máxima nos dados
min_date = min(khaolak_df['Date'].min(), competitors_df['Date'].min()).date()
max_date = max(khaolak_df['Date'].max(), competitors_df['Date'].max()).date()

# Criar inputs de data
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Data inicial", min_date, min_value=min_date, max_value=max_date)
with col2:
    end_date = st.date_input("Data final", max_date, min_value=min_date, max_value=max_date)

# Garantir que a data final não seja anterior à data inicial
if start_date > end_date:
    st.error('A data final deve ser posterior à data inicial.')
    st.stop()

# Filtrar os dados com base nas datas selecionadas
khaolak_filtered = khaolak_df[(khaolak_df['Date'].dt.date >= start_date) & (khaolak_df['Date'].dt.date <= end_date)]
competitors_filtered = competitors_df[(competitors_df['Date'].dt.date >= start_date) & (competitors_df['Date'].dt.date <= end_date)]

# Calcular as medianas
khaolak_median = khaolak_filtered.groupby('Date')['Price'].median().reset_index()
competitors_median = competitors_filtered.groupby('Date')['Price'].median().reset_index()


fig.add_trace(
    go.Scatter(
        x=khaolak_median['Date'],
        y=khaolak_median['Price'],
        mode='lines+markers',
        name='Median Khaolak',
        hovertemplate=(
            "<b>Data:</b> %{x|%Y-%m-%d}<br>"
            "<b>Mediana Khaolak:</b> %{y:.2f}<br>"
            "<b>Preço Médio Khaolak:</b> " + f"{khaolak_stats['mean']:.2f}" + "<br>"
            "<b>Preço Mínimo Khaolak:</b> " + f"{khaolak_stats['min']:.2f}" + "<br>"
            "<b>Preço Máximo Khaolak:</b> " + f"{khaolak_stats['max']:.2f}" + "<br>"
            "<extra></extra>"
        )
    )
)

fig.add_trace(
    go.Scatter(
        x=competitors_median['Date'],
        y=competitors_median['Price'],
        mode='lines+markers',
        name='Median competitors',
        hovertemplate=(
            "<b>Data:</b> %{x|%Y-%m-%d}<br>"
            "<b>Mediana Competidores:</b> %{y:.2f}<br>"
            "<b>Preço Médio Competidores:</b> " + f"{competitors_stats['mean']:.2f}" + "<br>"
            "<b>Preço Mínimo Competidores:</b> " + f"{competitors_stats['min']:.2f}" + "<br>"
            "<b>Preço Máximo Competidores:</b> " + f"{competitors_stats['max']:.2f}" + "<br>"
            "<b>Diferença Percentual:</b> " + f"{diff_percentage:.2f}%" + "<br>"
            "<extra></extra>"
        )
    )
)

# Configurar layout
fig.update_layout(
    title='Comparação de Medianas de Preços: Khaolak vs Competidores',
    xaxis_title='Data',
    yaxis_title='Preço (Mediana)',
    hovermode='x unified'
)

# Exibir o gráfico no Streamlit
st.plotly_chart(fig, use_container_width=True)

def calculate_daily_stats(df, date):
    day_data = df[df['Date'] == date]
    return {
        'date': date,
        'price': day_data['Price'].median(),
        'mean': day_data['Price'].mean(),
        'min': day_data['Price'].min(),
        'max': day_data['Price'].max(),
        'median': day_data['Price'].median()
    }

# Função para criar o relatório Excel
def create_excel_report(khaolak_df, competitors_df, start_date, end_date):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório Detalhado"

    # Cabeçalho
    headers = ['Data', 'Preço Khaolak', 'Média Khaolak', 'Mín Khaolak', 'Máx Khaolak', 'Mediana Khaolak',
               'Preço Competidores', 'Média Competidores', 'Mín Competidores', 'Máx Competidores', 'Mediana Competidores']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Dados diários
    date_range = pd.date_range(start=start_date, end=end_date)
    for date in date_range:
        khaolak_stats = calculate_daily_stats(khaolak_df, date)
        competitors_stats = calculate_daily_stats(competitors_df, date)
        
        row = [
            date.strftime('%Y-%m-%d'),
            khaolak_stats['price'],
            khaolak_stats['mean'],
            khaolak_stats['min'],
            khaolak_stats['max'],
            khaolak_stats['median'],
            competitors_stats['price'],
            competitors_stats['mean'],
            competitors_stats['min'],
            competitors_stats['max'],
            competitors_stats['median']
        ]
        ws.append(row)

    # Ajustar largura das colunas
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column_letter].width = adjusted_width

    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    return excel_buffer

# Exibir estatísticas gerais
st.subheader("Estatísticas Gerais para o Período Selecionado")
col1, col2 = st.columns(2)

with col1:
    st.write("Khaolak:")
    st.write(f"Preço Médio: {khaolak_stats['mean']:.2f}")
    st.write(f"Preço Mínimo: {khaolak_stats['min']:.2f}")
    st.write(f"Preço Máximo: {khaolak_stats['max']:.2f}")
    st.write(f"Preço Mediano: {khaolak_stats['median']:.2f}")

with col2:
    st.write("Competidores:")
    st.write(f"Preço Médio: {competitors_stats['mean']:.2f}")
    st.write(f"Preço Mínimo: {competitors_stats['min']:.2f}")
    st.write(f"Preço Máximo: {competitors_stats['max']:.2f}")
    st.write(f"Preço Mediano: {competitors_stats['median']:.2f}")

st.write(f"Diferença Percentual na Mediana: {diff_percentage:.2f}%")

excel_file = create_excel_report(khaolak_df, competitors_df, start_date, end_date)
st.download_button(
    label="Baixar Relatório Detalhado Excel",
    data=excel_file,
    file_name=f"relatorio_detalhado_{start_date}_a_{end_date}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)