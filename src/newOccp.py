import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import os
import io
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font

# Configuração da página
st.set_page_config(layout="wide")
st.title("Khaolak vs Competitors Dashboard")

global daily_occupancy, checkin_data
daily_occupancy = pd.DataFrame()
checkin_data = pd.DataFrame()

# Funções auxiliares
def clean_price(price):
    if pd.isna(price) or price == '':
        return None
    if isinstance(price, str):
        cleaned = ''.join(filter(str.isdigit, price))
        return float(cleaned) if cleaned else None
    return float(price)

def read_excel_files(directory):
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
                st.error(f"Error processing {filename}: {str(e)}")
    return competitors_dfs, khaolak_dfs

def calculate_stats(df, start_date, end_date):
    mask = (df['Date'] >= start_date) & (df['Date'] <= end_date)
    period_data = df.loc[mask]
    return {
        'mean': period_data['Price'].mean(),
        'min': period_data['Price'].min(),
        'max': period_data['Price'].max(),
        'median': period_data['Price'].median()
    }
    


def create_price_comparison_chart(khaolak_median, competitors_median, khaolak_stats, competitors_stats, diff_percentage):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
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
            name='Median Competitors',
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

    fig.update_layout(
        title='Comparação de Medianas de Preços: Khaolak vs Competidores',
        xaxis_title='Data',
        yaxis_title='Preço (Mediana)',
        hovermode='x unified'
    )

    return fig

def create_excel_report(khaolak_df, competitors_df, start_date, end_date):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        khaolak_df.to_excel(writer, sheet_name='Khaolak', index=False)
        competitors_df.to_excel(writer, sheet_name='Competitors', index=False)
    output.seek(0)
    return output

@st.cache_data
def load_price_data():
    directory = "C:/Users/ribei/Documents/RegiOtels/Dashboard-estatistica/DetailedPrices"
    competitors_dfs, khaolak_dfs = read_excel_files(directory)
    
    if not competitors_dfs and not khaolak_dfs:
        st.error("No valid files found. Check the directory and file names.")
        return None, None

    competitors_df = pd.concat(competitors_dfs, ignore_index=True) if competitors_dfs else pd.DataFrame()
    khaolak_df = pd.concat(khaolak_dfs, ignore_index=True) if khaolak_dfs else pd.DataFrame()

    khaolak_df['Date'] = pd.to_datetime(khaolak_df['Date'])
    competitors_df['Date'] = pd.to_datetime(competitors_df['Date'])

    return khaolak_df, competitors_df

@st.cache_data
# Funções para diferentes seções do dashboard
def price_comparison_section(khaolak_df, competitors_df):
    periods = {"1 Month": 30, "3 Months": 90, "6 Months": 180, "1 Year": 360}
    selected_period = st.selectbox("Select the viewing period:", list(periods.keys()))

    end_date = max(khaolak_df['Date'].max(), competitors_df['Date'].max())
    start_date = end_date - timedelta(days=periods[selected_period])

    khaolak_filtered = khaolak_df[(khaolak_df['Date'] >= start_date) & (khaolak_df['Date'] <= end_date)]
    competitors_filtered = competitors_df[(competitors_df['Date'] >= start_date) & (competitors_df['Date'] <= end_date)]

    khaolak_median = khaolak_filtered.groupby('Date')['Price'].median().reset_index()
    competitors_median = competitors_filtered.groupby('Date')['Price'].median().reset_index()

    khaolak_stats = calculate_stats(khaolak_filtered, start_date, end_date)
    competitors_stats = calculate_stats(competitors_filtered, start_date, end_date)

    diff_percentage = ((khaolak_stats['median'] - competitors_stats['median']) / competitors_stats['median']) * 100

    price_fig = create_price_comparison_chart(khaolak_median, competitors_median, khaolak_stats, competitors_stats, diff_percentage)
    st.plotly_chart(price_fig, use_container_width=True)

    return khaolak_filtered, competitors_filtered, khaolak_stats, competitors_stats, diff_percentage, start_date, end_date

def occupancy_section(daily_occupancy):
    st.subheader("Occupancy Chart")
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Start Date", 
                                   min_value=daily_occupancy.index.min(), 
                                   max_value=daily_occupancy.index.max(), 
                                   value=daily_occupancy.index.min())
    with col2:
        end_date = st.date_input("End Date", 
                                 min_value=start_date, 
                                 max_value=daily_occupancy.index.max(), 
                                 value=daily_occupancy.index.max())

    filtered_occupancy = daily_occupancy.loc[start_date:end_date]
    occupancy_fig = create_occupancy_chart(filtered_occupancy)
    st.plotly_chart(occupancy_fig, use_container_width=True)

    return filtered_occupancy, start_date, end_date

def hover_data_section(filtered_occupancy, checkin_data, start_date, end_date):
    st.subheader("Hover Data")
    hover_date = st.date_input("Select a date to view data", 
                               min_value=start_date, 
                               max_value=end_date, 
                               value=start_date)

    if hover_date not in filtered_occupancy.index:
        st.warning(f"No data available for {hover_date}. Showing the closest available date.")
        hover_date = min(filtered_occupancy.index, key=lambda x: abs(x - pd.Timestamp(hover_date)))

    hover_data = get_hover_data(hover_date, filtered_occupancy, checkin_data)
    st.dataframe(hover_data)

def statistics_section(khaolak_stats, competitors_stats, diff_percentage):
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


def read_checkin_files(main_directory):
    checkin_dfs = []
    for root, dirs, files in os.walk(main_directory):
        for filename in files:
            if filename.endswith(".xlsx"):
                file_path = os.path.join(root, filename)
                try:
                    df = pd.read_excel(file_path)
                    if all(col in df.columns for col in ['occupancy', 'checkin_date', 'price']):
                        df = df[df['occupancy'] == 2]
                        df['checkin_date'] = pd.to_datetime(df['checkin_date'], errors='coerce')
                        df['Hotel'] = os.path.basename(root)
                        checkin_dfs.append(df)
                except Exception as e:
                    st.error(f"Error processing file {filename}: {str(e)}")
    if not checkin_dfs:
        raise ValueError("No valid files found with required columns and occupancy equal to 2")
    return pd.concat(checkin_dfs, ignore_index=True)

def calculate_daily_occupancy(df):
    return df.groupby(['Hotel', 'checkin_date']).size().unstack(level=0).fillna(0)

def create_occupancy_chart(daily_occupancy):
    fig = go.Figure()
    for hotel in daily_occupancy.columns:
        fig.add_trace(go.Scatter(x=daily_occupancy.index, y=daily_occupancy[hotel], name=hotel, mode='none', fill='tonexty', stackgroup='one'))
    fig.update_layout(title='Daily Occupancy: Khaolak vs Competitors', xaxis_title='Timeline', yaxis_title='Available Rooms', hovermode='x unified', height=600)
    return fig



def get_hover_data(date, daily_occupancy, checkin_data):
    hover_data = []
    for hotel in daily_occupancy.columns:
        row = {
            'Hotel': hotel,
            'Date': date.strftime('%Y-%m-%d'),
            'Occupancy': daily_occupancy.loc[date, hotel] if date in daily_occupancy.index and hotel in daily_occupancy.columns else 'N/A'
        }
        hotel_data = checkin_data[(checkin_data['Hotel'] == hotel) & (checkin_data['checkin_date'] == date)]
        if not hotel_data.empty:
            row['Mean Price'] = f"{hotel_data['price'].mean():.2f}"
            row['Min Price'] = f"{hotel_data['price'].min():.2f}"
            row['Max Price'] = f"{hotel_data['price'].max():.2f}"
            row['Median Price'] = f"{hotel_data['price'].median():.2f}"
        else:
            row['Mean Price'] = row['Min Price'] = row['Max Price'] = row['Median Price'] = 'N/A'
        hover_data.append(row)
    return pd.DataFrame(hover_data)

main_directory = r"C:/Users/ribei/Documents/RegiOtels/Dashboard-estatistica/DashboardTHKHA"
try:
    checkin_data = read_checkin_files(main_directory)
    daily_occupancy = calculate_daily_occupancy(checkin_data)
except Exception as e:
    st.error(f"Error processing occupancy data: {str(e)}")
    daily_occupancy = pd.DataFrame()  # DataFrame vazio como fallback



def main():
    def main():

    # Carregar dados
    if 'khaolak_df' not in st.session_state or 'competitors_df' not in st.session_state:
        st.session_state.khaolak_df, st.session_state.competitors_df = load_price_data()
    
    if 'daily_occupancy' not in st.session_state or 'checkin_data' not in st.session_state:
        st.session_state.daily_occupancy, st.session_state.checkin_data = load_occupancy_data()

    if st.session_state.khaolak_df is None or st.session_state.competitors_df is None:
        st.error("Falha ao carregar dados de preços. Verifique os arquivos de origem.")
    return

    if st.session_state.daily_occupancy.empty:
        st.warning("Dados de ocupação não disponíveis. Verifique as fontes de dados.")
    return

    # Seção de comparação de preços
st.header("Comparação de Preços")
khaolak_filtered, competitors_filtered, khaolak_stats, competitors_stats, diff_percentage, price_start_date, price_end_date = price_comparison_section(st.session_state.khaolak_df, st.session_state.competitors_df)

    # Seção de ocupação
    st.header("Gráfico de Ocupação")
    
    # Seleção de intervalo de datas
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Data de Início", 
                                   min_value=st.session_state.daily_occupancy.index.min(), 
                                   max_value=st.session_state.daily_occupancy.index.max(), 
                                   value=st.session_state.daily_occupancy.index.min())
    with col2:
        end_date = st.date_input("Data de Fim", 
                                 min_value=start_date, 
                                 max_value=st.session_state.daily_occupancy.index.max(), 
                                 value=st.session_state.daily_occupancy.index.max())

    filtered_occupancy = st.session_state.daily_occupancy.loc[start_date:end_date]
    occupancy_fig = create_occupancy_chart(filtered_occupancy)
    st.plotly_chart(occupancy_fig, use_container_width=True)

    # Seção de dados do hover
    st.header("Dados Detalhados")
    hover_data_section(filtered_occupancy, st.session_state.checkin_data, start_date, end_date)

    # Seção de estatísticas
    st.header("Estatísticas do Período Selecionado")
    statistics_section(khaolak_stats, competitors_stats, diff_percentage)

    # Opção de download
    st.header("Download de Relatório")
    excel_file = create_excel_report(khaolak_filtered, competitors_filtered, price_start_date, price_end_date)
    st.download_button(
        label="Baixar Relatório Detalhado",
        data=excel_file,
        file_name=f"relatorio_detalhado_{price_start_date.date()}_a_{price_end_date.date()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()