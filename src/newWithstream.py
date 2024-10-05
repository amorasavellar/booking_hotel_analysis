import os
import glob
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from collections import Counter
from sklearn.linear_model import LinearRegression
import numpy as np
from datetime import datetime, timedelta

# Configuração da página
st.set_page_config(page_title="Khaolak Data Dashboard", layout="wide")


# Função para carregar e processar os dados
@st.cache_data
def load_data():
    # Path to the folder containing Excel files
    folder_path = r'C:/Users/ribei/Documents/RegioTels/Dashboard-estatistica/Dados'

# Find all Excel files in the specified folder
    files = glob.glob(os.path.join(folder_path, '*.xlsx'))

# List to store DataFrames
    df_list = []

    for file in files:
    # Extract the file name
        file_name = os.path.basename(file)

    # Split the file name by underscore
        parts = file_name.split('_')

    # Capture the part of the name that corresponds to the month
        month = parts[1]

    # Load the DataFrame from the Excel file
        df = pd.read_excel(file)

    # Add a column for the month, extracted from the file name
        df['Month'] = month

    # Add the DataFrame to the list
        df_list.append(df)

# Concatenate all DataFrames into a single DataFrame
    df_combined = pd.concat(df_list, ignore_index=True)

# Convert 'checkin_date' to datetime if it's not already
    df_combined['checkin_date'] = pd.to_datetime(df_combined['checkin_date'])

# Count occurrences of each check-in date
    checkin_counts = Counter(df_combined['checkin_date'])

# Convert the Counter to a DataFrame for easier visualization
    occupancy_df = pd.DataFrame.from_dict(checkin_counts, orient='index', columns=['Occupied_Rooms'])
    occupancy_df.index.name = 'Date'
    occupancy_df.reset_index(inplace=True)
    occupancy_df['Date'] = pd.to_datetime(occupancy_df['Date'])
    occupancy_df = occupancy_df.sort_values('Date')
        
    df_filtered = df_combined[df_combined['discount %'] >= 24]
        
    return df_combined, occupancy_df, df_filtered
    
# Carregando os dados
df_combined, occupancy_df, df_filtered = load_data()

st.sidebar.header("Filtros")
min_date = occupancy_df['Date'].min().date()
max_date = occupancy_df['Date'].max().date()
start_date = st.sidebar.date_input("Data de início", min_date)
end_date = st.sidebar.date_input("Data de fim", max_date)

# Filtro de hotel
hotels = df_combined['hotel_name'].unique()
selected_hotels = st.sidebar.multiselect("Selecione os hotéis", hotels, default=hotels)

# Aplicar filtros
mask = (occupancy_df['Date'].dt.date >= start_date) & (occupancy_df['Date'].dt.date <= end_date)
occupancy_df_filtered = occupancy_df.loc[mask]

mask = (df_filtered['checkin_date'].dt.date >= start_date) &\
       (df_filtered['checkin_date'].dt.date <= end_date) &\
       (df_filtered['hotel_name'].isin(selected_hotels))
df_filtered_filtered = df_filtered.loc[mask]

# Título do dashboard
st.title("Hotel Data Dashboard")

# Gráfico de ocupação
st.subheader("Hotel Occupancy")
fig_occupancy = go.Figure()
fig_occupancy.add_trace(go.Scatter(
    x=occupancy_df_filtered['Date'],
    y=occupancy_df_filtered['Occupied_Rooms'],
    mode='lines',
    name='Daily Occupancy'
))
fig_occupancy.update_layout(
    xaxis_title='Date',
    yaxis_title='Available Rooms',
    height=500
)
st.plotly_chart(fig_occupancy, use_container_width=True)

# Gráfico de distribuição de descontos
st.subheader("Discount Distribution")
fig_discount = go.Figure()
fig_discount.add_trace(go.Scatter(
    x=df_filtered_filtered['checkin_date'],
    y=df_filtered_filtered['discount %'],
    mode='markers',
    marker=dict(
        size=8,
        color=df_filtered_filtered['discount %'],
        colorscale='viridis',
        colorbar=dict(title="Discount %"),
        cmin=24,
        cmax=38,
    ),
    hovertemplate=
    "<b>Date:</b> %{x|%d/%m/%Y}<br>" +
    "<b>Discount:</b> %{y:.2f}%<br>" +
    "<b>Hotel:</b> %{customdata[0]}<br>" +
    "<b>Price:</b> R$ %{customdata[1]:.2f}<extra></extra>",
    customdata=df_filtered_filtered[['hotel_name', 'price']]
))
fig_discount.update_layout(
    xaxis_title="Check-in Date",
    yaxis_title="Discount Percentage",
    yaxis=dict(range=[24, 38]),
    height=500
)
st.plotly_chart(fig_discount, use_container_width=True)

# Gráfico de ocupação com regressão linear
st.subheader("Hotel Occupancy with Linear Regression")

# Preparar dados para regressão
X = (occupancy_df_filtered['Date'] - occupancy_df_filtered['Date'].min()).dt.days.values.reshape(-1, 1)
y = occupancy_df_filtered['Occupied_Rooms'].values

# Realizar regressão linear
model = LinearRegression()
model.fit(X, y)

# Gerar previsões
X_pred = np.array(range(X.min(), X.max() + 30)).reshape(-1, 1)  # Estender 30 dias para o futuro
y_pred = model.predict(X_pred)

# Converter X_pred de volta para datas
dates_pred = pd.date_range(start=occupancy_df_filtered['Date'].min(), periods=len(X_pred), freq='D')

# Criar o gráfico de ocupação com linha de regressão
fig_occupancy_regression = go.Figure()

# Adicionar a linha de ocupação
fig_occupancy_regression.add_trace(go.Scatter(
    x=occupancy_df_filtered['Date'],
    y=occupancy_df_filtered['Occupied_Rooms'],
    mode='lines',
    name='Daily Occupancy'
))

# Adicionar a linha de regressão
fig_occupancy_regression.add_trace(go.Scatter(
    x=dates_pred,
    y=y_pred,
    mode='lines',
    name='Linear Regression',
    line=dict(color='red', dash='dash')
))

fig_occupancy_regression.update_layout(
    xaxis_title='Date',
    yaxis_title='Available Rooms',
    height=500
)

st.plotly_chart(fig_occupancy_regression, use_container_width=True)

# Estatísticas de regressão
st.subheader("Regression Statistics")
st.write(f"Slope: {model.coef_[0]:.4f}")
st.write(f"Intercept: {model.intercept_:.4f}")
st.write(f"R-squared: {model.score(X, y):.4f}")