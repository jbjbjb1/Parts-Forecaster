import pandas as pd
from tkinter import filedialog, messagebox

def save_dataframes_to_excel(data, output_file):
    # Group by part number and month, sum the quantities
    data['Order Booked Month'] = data['Order Booked Date'].dt.to_period('M')
    grouped_data = data.groupby(['Item Number', 'Order Booked Month']).agg({'Net Qty Sold': 'sum'}).reset_index()

    # Convert 'Order Booked Month' back to a timestamp for consistency
    grouped_data['Order Booked Month'] = grouped_data['Order Booked Month'].dt.to_timestamp()

    # Create pivot table with 'Item Number' as rows and 'Order Booked Month' as columns
    pivot_table = grouped_data.pivot_table(
        index='Item Number',
        columns='Order Booked Month',
        values='Net Qty Sold',
        aggfunc='sum',
        fill_value=0
    )

    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Save the pivot table directly to the Excel file
        pivot_table.to_excel(writer, sheet_name='Processed Data', index=True)
        print("Pivot table saved to 'Processed Data' tab.")

def process_and_save_data(data):
    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save Processed Data",
        initialfile="parts_processed.xlsx"
    )
    
    if output_file:
        save_dataframes_to_excel(data, output_file)
        messagebox.showinfo("Success", f"Processed data has been saved to {output_file}")
    else:
        print("Save operation was canceled.")
