import pandas as pd
from prophet import Prophet
from prophet.plot import plot_plotly, plot_components_plotly

# Load data
df = pd.read_csv('https://raw.githubusercontent.com/facebook/prophet/main/examples/example_wp_log_peyton_manning.csv')

# Initialize and fit the model
m = Prophet()
m.fit(df)

# Create a future dataframe
future = m.make_future_dataframe(periods=365)

# Predict the future
forecast = m.predict(future)

# Plot using Plotly
plot_plotly(m, forecast).show()
plot_components_plotly(m, forecast).show()
