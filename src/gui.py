import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from process import process_and_save_data
from forecaster_ets import perform_ets_forecast  # Import the ETS method

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    # Load the Excel file
    try:
        data = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {str(e)}")
        return

    # Validate the columns
    expected_columns = {"Item Number", "Net Qty Sold", "Net Sales $", "Order Booked Date"}
    if set(data.columns) != expected_columns:
        messagebox.showerror("Error", "File does not have the required columns.")
        return

    # Process the data and save to Excel
    process_and_save_data(data)

def start_forecasting():
    from forecaster import process_data_for_prediction
    process_data_for_prediction()

def start_ets_forecasting():
    perform_ets_forecast()

# Set up the GUI
root = tk.Tk()
root.title("Forecasting Tool")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(padx=10, pady=10)

select_button = tk.Button(frame, text="(A) Select Excel File", command=select_file)
select_button.pack()

forecast_button = tk.Button(frame, text="(B) Prohpet Method", command=start_forecasting)
forecast_button.pack()

ets_button = tk.Button(frame, text="(B) ETS Method", command=start_ets_forecasting)  # New ETS button
ets_button.pack()

root.mainloop()
