import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from process import process_and_save_data
from forecaster import process_data_for_prediction

def select_file_for_processing():
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
    # Call the forecaster to process the saved Excel file and make predictions
    process_data_for_prediction()

# Set up the GUI
root = tk.Tk()
root.title("Forecasting Tool")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(padx=10, pady=10)

# Button to select the file for processing
process_button = tk.Button(frame, text="Select File to Process", command=select_file_for_processing)
process_button.pack(pady=5)

# Button to start the forecasting process
forecast_button = tk.Button(frame, text="Start Forecasting", command=start_forecasting)
forecast_button.pack(pady=5)

root.mainloop()
