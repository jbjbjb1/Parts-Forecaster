import pandas as pd
import matplotlib.pyplot as plt
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from tkinter import filedialog, messagebox

def perform_ets_forecast():
    # Prompt the user to select the intermediate Excel file
    input_file = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")],
        title="Select Processed Data File",
        initialfile="parts_processed.xlsx"
    )

    if not input_file:
        print("No file selected. Operation canceled.")
        return

    # Load the processed data (pivot table format)
    processed_data = pd.read_excel(input_file, sheet_name='Processed Data', index_col=0)

    # Sum the sales by month across all parts
    monthly_totals = processed_data.sum(axis=0)

    # Apply Holt-Winters Exponential Smoothing with quarterly seasonality
    if len(monthly_totals) < 8:
        messagebox.showwarning("Insufficient Data", "Not enough data for quarterly seasonal model. Using non-seasonal SES instead.")
        ses_model = SimpleExpSmoothing(monthly_totals).fit()
        forecast = ses_model.forecast(12)
    else:
        hw_model = ExponentialSmoothing(monthly_totals, seasonal='add', seasonal_periods=3).fit()
        forecast = hw_model.forecast(12)

    # Plot the actual and forecasted data
    plt.figure(figsize=(12, 8))
    plt.plot(monthly_totals.index, monthly_totals, marker='o', label='Actual Sales')
    plt.plot(forecast.index, forecast, marker='o', color='red', label='Holt-Winters Quarterly Forecast')
    plt.title('Actual Sales vs Holt-Winters Quarterly Forecast')
    plt.xlabel('Month')
    plt.ylabel('Total Quantity Sold')
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.legend()
    plt.tight_layout()
    plt.show()

    messagebox.showinfo("Forecast", "Holt-Winters quarterly forecast has been plotted.")
