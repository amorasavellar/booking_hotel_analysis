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

st.set_page_config(page_title="Price and Occupancy Comparison", layout="wide")

def clean_price(price):
    if pd.isna(price) or price == '':
        return None
    if isinstance(price, str):
        cleaned = ''.join(filter(str.isdigit, price))
        return float(cleaned) if cleaned else None
    return float(price)

directory = "C:/Users/ribei/Documents/RegiOtels/Dashboard-estatistica/DetailedPrices"
competitors_dfs = []
khaolak_dfs = []

for filename in os.listdir(directory):
    if filename.endswith(".xlsx") and "detailed_prices" in filename:
        file_path = os.path.join(directory, filename)
        try:
            df = pd.read_excel(file_path)
            df['Hotel'] = filename.split('_')[0]
            df['Price'] = df['Price'].apply(clean_price)
            df_cleaned = df.dropna(subset=['Price'])
            if "khaolak" in filename.lower():
                khaolak_dfs.append(df_cleaned)
            else:
                competitors_dfs.append(df_cleaned)
        except Exception as e:
            st.error(f"Erro ao processar {filename}: {str(e)}")

if not competitors_dfs and not khaolak_dfs:
    st.error("Nenhum arquivo válido encontrado. Verifique o diretório e os nomes dos arquivos.")
    st.stop()

competitors_df = pd.concat(competitors_dfs, ignore_index=True) if competitors_dfs else pd.DataFrame()
khaolak_df = pd.concat(khaolak_dfs, ignore_index=True) if khaolak_dfs else pd.DataFrame()

def calculate_stats(df, start_date, end_date):
    mask = (df['Date'] >= start_date) & (df['Date'] <= end_date)
    period_data = df.loc[mask]
    return {
        'mean': period_data['Price'].mean(),
        'min': period_data['Price'].min(),
        'max': period_data['Price'].max(),
        'median': period_data['Price'].median()
    }
    
st.title("Khaolak vs Competitors")

periods = {
    "1 Month": 30,
    "3 Months": 90,
    "6 Months": 180,
    "1 Year": 360,
}
selected_period = st.selectbox("Select the viewing period:", list(periods.keys()))

khaolak_df['Date'] = pd.to_datetime(khaolak_df['Date'])
competitors_df['Date'] = pd.to_datetime(competitors_df['Date'])

end_date = max(khaolak_df['Date'].max(), competitors_df['Date'].max())
start_date = end_date - timedelta(days=periods[selected_period])

khaolak_filtered = khaolak_df[(khaolak_df['Date'] >= start_date) & (khaolak_df['Date'] <= end_date)]
competitors_filtered = competitors_df[(competitors_df['Date'] >= start_date) & (competitors_df['Date'] <= end_date)]

khaolak_median = khaolak_filtered.groupby('Date')['Price'].median().reset_index()
competitors_median = competitors_filtered.groupby('Date')['Price'].median().reset_index()

khaolak_stats = calculate_stats(khaolak_filtered, start_date, end_date)
competitors_stats = calculate_stats(competitors_filtered, start_date, end_date)

diff_percentage = ((khaolak_stats['median'] - competitors_stats['median']) / competitors_stats['median']) * 100

# Criar o gráfico de comparação de preços
fig = make_subplots(specs=[[{"secondary_y": True}]])

fig.add_trace(
    go.Scatter(
        x=khaolak_median['Date'],
        y=khaolak_median['Price'],
        mode='lines+markers',
        name='Median Khaolak',
        hovertemplate=(
            "<b>Date:</b> %{x|%Y-%m-%d}<br>"
            "<b>Median Khaolak:</b> %{y:.2f}<br>"
            "<b>Mean Price Khaolak:</b> " + f"{khaolak_stats['mean']:.2f}" + "<br>"
            "<b>Minimum Price Khaolak:</b> " + f"{khaolak_stats['min']:.2f}" + "<br>"
            "<b>Maximum Price Khaolak:</b> " + f"{khaolak_stats['max']:.2f}" + "<br>"
            "<extra></extra>"
        )
    )
)

fig.add_trace(
    go.Scatter(
        x=competitors_median['Date'],
        y=competitors_median['Price'],
        mode='lines+markers',
        name='Median Competitors',
        hovertemplate=(
            "<b>Date:</b> %{x|%Y-%m-%d}<br>"
            "<b>Median Competitors:</b> %{y:.2f}<br>"
            "<b>Mean Price Competitors:</b> " + f"{competitors_stats['mean']:.2f}" + "<br>"
            "<b>Minimum Price Competitors:</b> " + f"{competitors_stats['min']:.2f}" + "<br>"
            "<b>Maximum Price Competitors:</b> " + f"{competitors_stats['max']:.2f}" + "<br>"
            "<b>Percentage Difference:</b> " + f"{diff_percentage:.2f}%" + "<br>"
            "<extra></extra>"
        )
    )
)

fig.update_layout(
    title='Comparison of Median Prices: Khaolak vs Competitors',
    xaxis_title='Timeline',
    yaxis_title='Price (Median)',
    hovermode='x unified'
)

# Exibir o gráfico no Streamlit
st.plotly_chart(fig, use_container_width=True)

def read_checkin_files(main_directory):
    checkin_dfs = []
    
    for root, dirs, files in os.walk(main_directory):
        for filename in files:
            if filename.endswith(".xlsx"):
                file_path = os.path.join(root, filename)
                try:
                    df = pd.read_excel(file_path)
                    required_columns = ['occupancy', 'checkin_date', 'price']
                    if not all(col in df.columns for col in required_columns):
                        continue
                    
                    df = df[df['occupancy'] == 2]
                    df['checkin_date'] = pd.to_datetime(df['checkin_date'], errors='coerce')
                    
                    # Usar o nome da pasta como nome do hotel
                    hotel_name = os.path.basename(root)
                    df['Hotel'] = hotel_name
                    
                    checkin_dfs.append(df)
                except Exception as e:
                    st.error(f"Erro ao processar o arquivo {filename}: {str(e)}")
    
    if not checkin_dfs:
        raise ValueError("Nenhum arquivo válido encontrado com as colunas necessárias e occupancy igual a 2")
    
    return pd.concat(checkin_dfs, ignore_index=True)

def calculate_daily_occupancy(df):
    return df.groupby(['Hotel', 'checkin_date']).size().unstack(level=0).fillna(0)

main_directory = r"C:/Users/ribei/Documents/RegiOtels/Dashboard-estatistica/DashboardTHKHA"
try:
    checkin_data = read_checkin_files(main_directory)
    daily_occupancy = calculate_daily_occupancy(checkin_data)
except Exception as e:
    st.error(f"Erro ao processar dados de ocupação: {str(e)}")
    st.stop()
    
st.subheader("Occupancy Chart")

# Seletor de período para o gráfico de ocupação
occupancy_periods = {
    "1 Month": 30,
    "3 Months": 90,
    "6 Months": 180,
    "Entire Period": None  # Adicionamos uma opção para ver todo o período
}
selected_occupancy_period = st.selectbox("Select the viewing period:", list(occupancy_periods.keys()))

# Calcular as datas de início e fim para o gráfico de ocupação
occupancy_end_date = daily_occupancy.index.max()
if occupancy_periods[selected_occupancy_period] is not None:
    occupancy_start_date = occupancy_end_date - timedelta(days=occupancy_periods[selected_occupancy_period])
else:
    occupancy_start_date = daily_occupancy.index.min()

# Filtrar os dados de ocupação para o período selecionado
filtered_occupancy = daily_occupancy.loc[occupancy_start_date:occupancy_end_date]

def create_occupancy_chart(daily_occupancy, start_date, end_date):
    fig = go.Figure()
    
    sorted_columns = daily_occupancy.sum().sort_values().index
    
    for hotel in sorted_columns:
        fig.add_trace(go.Scatter(
            x=daily_occupancy.index,
            y=daily_occupancy[hotel],
            name=hotel,
            mode='none',
            fill='tonexty',
            stackgroup='one'
        ))
    
    fig.update_layout(
        title='Daily Occupancy: Khaolak vs Competitors',
        xaxis_title='Timeline',
        yaxis_title='Available Rooms',
        hovermode='x unified',
        legend_title='Hotéis',
        showlegend=True
    )
    
    return fig

# Criar e exibir o gráfico de ocupação com os dados filtrados
if not filtered_occupancy.empty:
    occupancy_fig = create_occupancy_chart(filtered_occupancy, occupancy_start_date, occupancy_end_date)
    st.plotly_chart(occupancy_fig, use_container_width=True)
else:
    st.warning("Não foi possível criar o gráfico de ocupação devido à falta de dados válidos para o período selecionado")
    
# Exibir estatísticas gerais
st.subheader("Statistics for the Selected Period")
col1, col2 = st.columns(2)

with col1:
    st.write("Khaolak:")
    st.write(f"Mean Price: {khaolak_stats['mean']:.2f}")
    st.write(f"Minimum Price: {khaolak_stats['min']:.2f}")
    st.write(f"Maximum Price: {khaolak_stats['max']:.2f}")
    st.write(f"Median Price: {khaolak_stats['median']:.2f}")

with col2:
    st.write("Competitors:")
    st.write(f"Mean Price: {competitors_stats['mean']:.2f}")
    st.write(f"Minimum Price: {competitors_stats['min']:.2f}")
    st.write(f"Maximum Price: {competitors_stats['max']:.2f}")
    st.write(f"Median Price: {competitors_stats['median']:.2f}")

st.write(f"Percentage Difference in Median: {diff_percentage:.2f}%")

# Função para criar o relatório Excel
def create_excel_report(khaolak_df, competitors_df, start_date, end_date):
    wb = Workbook()
    ws = wb.active
    ws.title = "Detailed Report"

    headers = ['Data', 'Preço Khaolak', 'Média Khaolak', 'Mín Khaolak', 'Máx Khaolak', 'Mediana Khaolak',
               'Preço Competidores', 'Média Competidores', 'Mín Competidores', 'Máx Competidores', 'Mediana Competidores']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    date_range = pd.date_range(start=start_date, end=end_date)
    for date in date_range:
        khaolak_stats = calculate_stats(khaolak_df[khaolak_df['Date'] == date], date, date)
        competitors_stats = calculate_stats(competitors_df[competitors_df['Date'] == date], date, date)
        
        row = [
            date.strftime('%Y-%m-%d'),
            khaolak_stats['median'],
            khaolak_stats['mean'],
            khaolak_stats['min'],
            khaolak_stats['max'],
            khaolak_stats['median'],
            competitors_stats['median'],
            competitors_stats['mean'],
            competitors_stats['min'],
            competitors_stats['max'],
            competitors_stats['median']
        ]
        ws.append(row)

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

# Criar e oferecer download do relatório Excel
excel_file = create_excel_report(khaolak_df, competitors_df, start_date, end_date)
st.download_button(
    label="Download Detailed Report of the Selected Period in Excel",
    data=excel_file,
    file_name=f"detailed_report_{start_date.date()}_a_{end_date.date()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)