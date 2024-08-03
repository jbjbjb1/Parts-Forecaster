import pandas as pd
from prophet import Prophet
from tkinter import messagebox, filedialog
import numpy as np
import matplotlib.pyplot as plt

def predict_sales(item, item_data):
    # Reset the index to access the date columns
    item_data = item_data.reset_index()

    # Melt the pivot table back to long format
    melted_data = item_data.melt(id_vars=['Item Number'], var_name='ds', value_name='y')
    
    # Convert 'ds' to datetime
    melted_data['ds'] = pd.to_datetime(melted_data['ds'])

    # Filter out zero values
    melted_data = melted_data[melted_data['y'] > 0]

    unique_dates = melted_data['ds'].nunique()

    if unique_dates <= 3:
        # Predict 50% of what was done in the last 12 months
        last_12_months = melted_data[melted_data['ds'] >= pd.Timestamp.today() - pd.DateOffset(years=1)]
        total_sales = last_12_months['y'].sum()
        predicted_sales = np.floor(0.5 * total_sales)  # Round down to the nearest integer

        if not last_12_months.empty:
            # Assign prediction to the same date one year in the future
            last_sale_date = last_12_months['ds'].max()

            # Ensure that last_sale_date is a Timestamp
            if not pd.isna(last_sale_date):
                future_date = last_sale_date + pd.DateOffset(years=1)

                # Check if the future date falls within the next 12 months from today
                if future_date <= pd.Timestamp.today() + pd.DateOffset(years=1):
                    predicted_df = pd.DataFrame({'ds': [future_date], 'y': [predicted_sales], 'Item Number': [item]})
                    print(f"Predicting {predicted_sales} on {future_date}.")
                else:
                    predicted_df = pd.DataFrame(columns=['ds', 'y', 'Item Number'])  # No prediction if outside the next 12 months
                    print(f"No prediction, future date {future_date} is outside the next 12 months.")
            else:
                predicted_df = pd.DataFrame(columns=['ds', 'y', 'Item Number'])  # No prediction if no valid last sale date
                print(f"No valid last sale date for item {item}.")
        else:
            # If there were no sales in the last year, predict zero for the future
            future_date = pd.Timestamp.today() + pd.DateOffset(years=1)
            predicted_df = pd.DataFrame({'ds': [future_date], 'y': [0], 'Item Number': [item]})
            print(f"No sales in the last year for item {item}. Setting future sales to zero.")
    else:
        # Use Prophet forecaster with monthly frequency
        model = Prophet()
        model.fit(melted_data[['ds', 'y']])
        future_df = model.make_future_dataframe(periods=12, freq='M')  # Adjust to monthly periods
        forecast = model.predict(future_df)
        forecast['yhat'] = forecast['yhat'].apply(lambda x: max(0, np.floor(x)))  # Round down and replace negative values with 0
        predicted_df = forecast[['ds', 'yhat']].rename(columns={'yhat': 'y'})
        predicted_df['Item Number'] = item
        print("Using Prophet for prediction.")

    # Ensure the 'ds' column is month-based for the pivot table, only if the DataFrame is not empty
    if not predicted_df.empty:
        predicted_df['ds'] = predicted_df['ds'].dt.to_period('M').dt.to_timestamp()

    return predicted_df

def plot_predictions(pivot_data, prophet_forecast):
    # Sum the predicted sales for each month
    monthly_totals = pivot_data.sum(axis=0)
    
    # Plot the sum of each month
    plt.figure(figsize=(12, 8))
    monthly_totals.plot(kind='line', marker='o', label='Sum of Individual Sales Predictions')
    
    # Plot the Prophet forecast for total monthly sales
    plt.plot(prophet_forecast['ds'], prophet_forecast['yhat'], marker='o', color='red', label='Prophet Forecast of Total Sales')
    
    plt.title('Total Predicted Sales by Month')
    plt.xlabel('Month')
    plt.ylabel('Total Quantity Sold')
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.legend()
    plt.tight_layout()
    plt.show()

def process_data_for_prediction():
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

    # Predict sales for each part and store the results
    predicted_data = []
    for item in processed_data.index:
        item_data = processed_data.loc[[item]]
        predicted_data.append(predict_sales(item, item_data))

    # Concatenate all the predicted data
    full_predicted_df = pd.concat(predicted_data)

    # Create pivot table from the predicted data
    pivot_data = full_predicted_df.pivot_table(
        index='Item Number',
        columns='ds',
        values='y',
        aggfunc='sum',
        fill_value=0
    )

    # Prophet forecast on aggregated monthly sales
    monthly_totals_df = pivot_data.sum(axis=0).reset_index()
    monthly_totals_df.columns = ['ds', 'y']
    
    # Apply caps or growth constraints here
    monthly_totals_df['cap'] = monthly_totals_df['y'].max()  # Set a cap based on the maximum observed value
    prophet_model = Prophet(growth='logistic', yearly_seasonality=False)  # Disable yearly seasonality if unnecessary
    prophet_model.fit(monthly_totals_df)
    future = prophet_model.make_future_dataframe(periods=12, freq='M')
    future['cap'] = monthly_totals_df['cap'].max()  # Ensure future cap is set
    prophet_forecast = prophet_model.predict(future)

    final_output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save Final Predictions",
        initialfile="final_predictions.xlsx"
    )

    if final_output_file:
        with pd.ExcelWriter(final_output_file, engine='xlsxwriter') as writer:
            pivot_data.to_excel(writer, sheet_name='Predictions', index=True)
        print(f"Final predictions saved to {final_output_file}.")
        messagebox.showinfo("Success", f"Final predictions have been saved to {final_output_file}")
        
        # Plot the sum of predictions for each month and the Prophet forecast
        plot_predictions(pivot_data, prophet_forecast)
        
    else:
        print("Save operation was canceled.")
