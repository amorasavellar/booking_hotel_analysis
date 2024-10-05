import os
import glob
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from collections import Counter
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import PolynomialFeatures
from plotly.subplots import make_subplots



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


fig = make_subplots(specs=[[{"secondary_y": True}]])

# Create the occupancy graph
fig_occupancy = go.Figure()

# Add the occupancy line
fig_occupancy.add_trace(go.Scatter(
    x=occupancy_df['Date'],
    y=occupancy_df['Occupied_Rooms'],
    mode='lines',
    name='Daily Occupancy'
))

# Configure the layout for occupancy graph
fig_occupancy.update_layout(
    title='Hotel Occupancy - Interactive Zoom',
    xaxis_title='Date',
    yaxis_title='Available Rooms',
    hovermode='x unified',
    xaxis=dict(
        rangeselector=dict(
            buttons=list([
                dict(count=1, label="1m", step="month", stepmode="backward"),
                dict(count=3, label="3m", step="month", stepmode="backward"),
                dict(count=6, label="6m", step="month", stepmode="backward"),
                dict(count=1, label="1y", step="year", stepmode="backward"),
                dict(step="all")
            ])
        ),
        rangeslider=dict(visible=True),
        type="date"
    )
)

# Show the occupancy graph
fig_occupancy.show()

# Create the discount distribution graph
fig_discount = go.Figure()

# Filter out rows where discount % is less than 24
df_filtered = df_combined[df_combined['discount %'] >= 24]

fig_discount.add_trace(go.Scatter(
    x=df_filtered['checkin_date'],
    y=df_filtered['discount %'],
    mode='markers',
    marker=dict(
        size=8,
        color=df_filtered['discount %'],
        colorscale='viridis',
        colorbar=dict(title="Discount %"),
        cmin=24,
        cmax=38,
    ),
    hovertemplate=
    "<b>Data:</b> %{x|%d/%m/%Y}<br>" +
    "<b>Desconto:</b> %{y:.2f}%<br>" +
    "<b>Hotel:</b> %{customdata[0]}<br>" +
    "<b>Preço:</b> R$ %{customdata[1]:.2f}<extra></extra>",
    customdata=df_filtered[['hotel_name', 'price']]
))

# Update layout for discount graph
fig_discount.update_layout(
    title='Discounts Distribution Over Time',
    xaxis_title="Check-in Date",
    yaxis_title="Discount Percentage",
    yaxis=dict(range=[24, 38]),
    xaxis=dict(
        rangeselector=dict(
            buttons=list([
                dict(count=1, label="1m", step="month", stepmode="backward"),
                dict(count=6, label="6m", step="month", stepmode="backward"),
                dict(count=1, label="1y", step="year", stepmode="backward"),
                dict(step="all")
            ])
        ),
        rangeslider=dict(visible=True, thickness=0.05),
        type="date"
    ),
)

# Show the discount distribution graph
fig_discount.show()

# Print summary statistics
print("Occupancy Statistics:")
print(occupancy_df.describe())
print("\nDiscount Distribution Statistics:")
print(df_filtered.groupby(df_filtered['checkin_date'].dt.to_period('D'))['discount %'].describe())


# Prepare data for regression
X = (occupancy_df['Date'] - occupancy_df['Date'].min()).dt.days.values.reshape(-1, 1)
y = occupancy_df['Occupied_Rooms'].values

# Perform linear regression
model = LinearRegression()
model.fit(X, y)

# Generate predictions
X_pred = np.array(range(X.min(), X.max() + 30)).reshape(-1, 1)  # Extend 30 days into the future
y_pred = model.predict(X_pred)

# Convert X_pred back to dates
dates_pred = pd.date_range(start=occupancy_df['Date'].min(), periods=len(X_pred), freq='D')

# Create the occupancy graph with regression line
fig_occupancy = go.Figure()

# Add the occupancy line
fig_occupancy.add_trace(go.Scatter(
    x=occupancy_df['Date'],
    y=occupancy_df['Occupied_Rooms'],
    mode='lines',
    name='Daily Occupancy'
))

# Add the regression line
fig_occupancy.add_trace(go.Scatter(
    x=dates_pred,
    y=y_pred,
    mode='lines',
    name='Linear Regression',
    line=dict(color='red', dash='dash')
))

# Configure the layout for occupancy graph
fig_occupancy.update_layout(
    title='Hotel Occupancy with Linear Regression',
    xaxis_title='Date',
    yaxis_title='Available Rooms',
    hovermode='x unified',
    xaxis=dict(
        rangeselector=dict(
            buttons=list([
                dict(count=1, label="1m", step="month", stepmode="backward"),
                dict(count=3, label="3m", step="month", stepmode="backward"),
                dict(count=6, label="6m", step="month", stepmode="backward"),
                dict(count=1, label="1y", step="year", stepmode="backward"),
                dict(step="all")
            ])
        ),
        rangeslider=dict(visible=True),
        type="date"
    )
)

# Show the occupancy graph
fig_occupancy.show()

# Print regression statistics
print("Linear Regression Statistics:")
print(f"Slope: {model.coef_[0]:.4f}")
print(f"Intercept: {model.intercept_:.4f}")
print(f"R-squared: {model.score(X, y):.4f}")


fig.add_trace(
    go.Scatter(
        x=occupancy_df['Date'],
        y=occupancy_df['Occupied_Rooms'],
        name="Available Rooms",
        mode='lines',
        line=dict(color="blue")
    ),
    secondary_y=False,
)

# Adicionar o gráfico de dispersão de descontos no eixo secundário
fig.add_trace(
    go.Scatter(
        x=df_filtered['checkin_date'],
        y=df_filtered['discount %'],
        mode='markers',
        name="Discount",
        marker=dict(
            size=8,
            color=df_filtered['discount %'],
            colorscale='viridis',
            colorbar=dict(title="Desconto %", x=1.1),  # Movido para a direita
            cmin=24,
            cmax=38,
        ),
        hovertemplate=
        "<b>Data:</b> %{x|%d/%m/%Y}<br>" +
        "<b>Desconto:</b> %{y:.2f}%<br>" +
        "<b>Hotel:</b> %{customdata[0]}<br>" +
        "<b>Preço:</b> R$ %{customdata[1]:.2f}<extra></extra>",
        customdata=df_filtered[['hotel_name', 'price']]
    ),
    secondary_y=True,
)

# Configurar os títulos dos eixos
fig.update_xaxes(title_text="Data")
fig.update_yaxes(title_text="<b>Available Rooms</b>", secondary_y=False, title_standoff=15)
fig.update_yaxes(title_text="<b>Porcentagem de Desconto</b>", secondary_y=True, title_standoff=15)

# Atualizar o layout
fig.update_layout(
    title='Available Rooms Over Time vs. Discounts Distribution Over Time',
    hovermode="x unified",
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    ),
    margin=dict(l=50, r=50, t=100, b=50),
    xaxis=dict(
        rangeselector=dict(
            buttons=list([
                dict(count=1, label="1m", step="month", stepmode="backward"),
                dict(count=3, label="3m", step="month", stepmode="backward"),
                dict(count=6, label="6m", step="month", stepmode="backward"),
                dict(count=1, label="1y", step="year", stepmode="backward"),
                dict(step="all")
            ])
        ),
        rangeslider=dict(visible=True),
        type="date"
    )
)

# Mostrar o gráfico
fig.show()