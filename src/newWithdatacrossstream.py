import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from collections import Counter
from sklearn.linear_model import LinearRegression
import os
import glob
from io import BytesIO

# Streamlit page configuration
st.set_page_config(page_title="Hotel Analytics Dashboard", layout="wide")
st.title("Hotel Analytics Dashboard")

# Function to load data
@st.cache_data
def load_data(folder_path):
    files = glob.glob(os.path.join(folder_path, '*.xlsx'))
    df_list = []
    for file in files:
        file_name = os.path.basename(file)
        parts = file_name.split('_')
        month = parts[1]
        df = pd.read_excel(file)
        df['Month'] = month
        df_list.append(df)
    df_combined = pd.concat(df_list, ignore_index=True)
    df_combined['checkin_date'] = pd.to_datetime(df_combined['checkin_date'])
    return df_combined

# Load data
folder_path = r'C:/Users/ribei/Documents/RegioTels/Dashboard-estatistica/Dados'
df_combined = load_data(folder_path)

# Create occupancy DataFrame
checkin_counts = Counter(df_combined['checkin_date'])
occupancy_df = pd.DataFrame.from_dict(checkin_counts, orient='index', columns=['Occupied_Rooms'])
occupancy_df.index.name = 'Date'
occupancy_df.reset_index(inplace=True)
occupancy_df['Date'] = pd.to_datetime(occupancy_df['Date'])
occupancy_df = occupancy_df.sort_values('Date')

# Sidebar for period selection
st.sidebar.header("Settings")
date_range = st.sidebar.selectbox(
    "Select analysis period:",
    ["1 month", "3 months", "6 months", "All data"]
)

# Function to filter data based on selected period
def filter_data(df, date_column, period):
    end_date = df[date_column].max()
    if period == "1 month":
        start_date = end_date - pd.Timedelta(days=30)
    elif period == "3 months":
        start_date = end_date - pd.Timedelta(days=90)
    elif period == "6 months":
        start_date = end_date - pd.Timedelta(days=180)
    else:
        return df
    return df[(df[date_column] >= start_date) & (df[date_column] <= end_date)]

# Filter data
filtered_occupancy_df = filter_data(occupancy_df, 'Date', date_range)
filtered_df_combined = filter_data(df_combined, 'checkin_date', date_range)


# Occupancy graph
st.header("Hotel Occupancy")
fig_occupancy = go.Figure()
fig_occupancy.add_trace(go.Scatter(
    x=filtered_occupancy_df['Date'],
    y=filtered_occupancy_df['Occupied_Rooms'],
    mode='lines',
    name='Daily Occupancy'
))

# Linear regression
X = (filtered_occupancy_df['Date'] - filtered_occupancy_df['Date'].min()).dt.days.values.reshape(-1, 1)
y = filtered_occupancy_df['Occupied_Rooms'].values
model = LinearRegression()
model.fit(X, y)
X_pred = np.array(range(X.min(), X.max() + 30)).reshape(-1, 1)
y_pred = model.predict(X_pred)
dates_pred = pd.date_range(start=filtered_occupancy_df['Date'].min(), periods=len(X_pred), freq='D')

fig_occupancy.add_trace(go.Scatter(
    x=dates_pred,
    y=y_pred,
    mode='lines',
    name='Linear Regression',
    line=dict(color='red', dash='dash')
))

fig_occupancy.update_layout(
    xaxis_title='Date',
    yaxis_title='Available Rooms',
    yaxis=dict(rangemode='nonnegative', zeroline=True, zerolinewidth=2, zerolinecolor='LightGray'),
    hovermode='x unified'
)

# Ensure y-axis starts at 1
y_max = max(filtered_occupancy_df['Occupied_Rooms'].max(), y_pred.max())
fig_occupancy.update_yaxes(range=[1, y_max * 1.1])  # Start at 1 and add 10% padding at the top

st.plotly_chart(fig_occupancy, use_container_width=True)

# Função para criar o texto do hover
def create_hover_text(row):
    return f"<b>Date:</b> {row['checkin_date'].strftime('%d/%m/%Y')}<br>" + \
           f"<b>Discount:</b> {row['discount %']:.2f}%<br>" + \
           f"<b>Hotel:</b> {row['hotel_name']}<br>" + \
           f"<b>Price:</b> R$ {row['price']:.2f}"

# Filtrar dados para excluir descontos de 0%
filtered_df_combined_nonzero = filtered_df_combined[filtered_df_combined['discount %'] > 0]

# Gráfico de distribuição de descontos
st.header("Discount Distribution")
fig_discount = go.Figure()

fig_discount.add_trace(go.Scatter(
    x=filtered_df_combined_nonzero['checkin_date'],
    y=filtered_df_combined_nonzero['discount %'],
    mode='markers',
    marker=dict(
        size=8,
        color=filtered_df_combined_nonzero['discount %'],
        colorscale='viridis',
        colorbar=dict(title="Discount %"),
        cmin=filtered_df_combined_nonzero['discount %'].min(),
        cmax=filtered_df_combined_nonzero['discount %'].max(),
    ),
    text=filtered_df_combined_nonzero.apply(create_hover_text, axis=1),
    hoverinfo='text',
    hoverlabel=dict(namelength=-1)
))

fig_discount.update_layout(
    xaxis_title="Check-in Date",
    yaxis_title="Discount Percentage",
    yaxis=dict(range=[filtered_df_combined_nonzero['discount %'].min() * 0.95, 
                      filtered_df_combined_nonzero['discount %'].max() * 1.05]),
    height=500,
    hovermode="closest"
)

st.plotly_chart(fig_discount, use_container_width=True)

# Gráfico combinado de ocupação e descontos
st.header("Occupancy vs Discounts")
fig_combined = make_subplots(specs=[[{"secondary_y": True}]])

# Adicionar traço de ocupação
fig_combined.add_trace(
    go.Scatter(
        x=filtered_occupancy_df['Date'],
        y=filtered_occupancy_df['Occupied_Rooms'],
        name="Available Rooms",
        mode='lines',
        line=dict(color="blue", width=2)
    ),
    secondary_y=False,
)

# Adicionar traço de desconto
fig_combined.add_trace(
    go.Scatter(
        x=filtered_df_combined_nonzero['checkin_date'],
        y=filtered_df_combined_nonzero['discount %'],
        mode='markers',
        name="Discount",
        marker=dict(
            size=8,
            color=filtered_df_combined_nonzero['discount %'],
            colorscale='viridis',
            colorbar=dict(title="Discount %",),
            cmin=filtered_df_combined_nonzero['discount %'].min(),
            cmax=filtered_df_combined_nonzero['discount %'].max(),
        ),
        text=filtered_df_combined_nonzero.apply(create_hover_text, axis=1),
        hoverinfo='text',
        hoverlabel=dict(namelength=-1)
    ),
    secondary_y=True,
)

fig_combined.update_layout(
    xaxis_title="Date",
    yaxis_title="Occupied Rooms",
    yaxis2_title="Discount Percentage",
    hovermode="closest",
    height=600,
    margin=dict(r=70, t=60),  # Aumentado o topo para dar mais espaço
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1,  # Movido para cima
        xanchor="center",  # Centralizado
        x=0.5  # Centralizado
    )
)

# Ajustar intervalos do eixo y
fig_combined.update_yaxes(range=[1, max(filtered_occupancy_df['Occupied_Rooms']) * 1.1], secondary_y=False)
fig_combined.update_yaxes(
    range=[filtered_df_combined_nonzero['discount %'].min() * 0.95, 
           filtered_df_combined_nonzero['discount %'].max() * 1.05], 
    secondary_y=True
)

st.plotly_chart(fig_combined, use_container_width=True)

# Summary statistics
st.header("Summary Statistics")
col1, col2 = st.columns(2)

with col1:
    st.subheader("Occupancy")
    st.write(filtered_occupancy_df['Occupied_Rooms'].describe())

with col2:
    st.subheader("Discount Distribution")
    discount_stats = filtered_df_combined.groupby(filtered_df_combined['checkin_date'].dt.to_period('D'))['discount %'].describe()
    st.write(discount_stats)
    
    # Add download button for discount data
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Sheet1')
        processed_data = output.getvalue()
        return processed_data

    excel_file = to_excel(discount_stats)
    st.download_button(
        label="Download Discount Data as Excel",
        data=excel_file,
        file_name="discount_statistics.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Footer note
st.markdown("---")
st.markdown("Dashboard created with Streamlit and Plotly")