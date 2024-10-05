import pandas as pd
from datetime import datetime, timedelta
import os
import re

def remove_apostrophe(x):
    return x.lstrip("'") if isinstance(x, str) else x

def process_hotel_file(file_path):
    try:
        print(f"\nProcessing file: {file_path}")
        
        # Read all sheets in the Excel file
        xls = pd.ExcelFile(file_path)
        print(f"Sheets in the file: {xls.sheet_names}")
        
        # Try to find a sheet with the required columns
        required_columns = ['checkin_date', 'price', 'occupancy', 'breakfast_included', 'hotel_name', 'refundable', 'name', 'type']
        for sheet_name in xls.sheet_names:
            print(f"Checking sheet: {sheet_name}")
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"Columns in this sheet: {df.columns.tolist()}")
            
            if all(col in df.columns for col in required_columns):
                print("Found sheet with required columns")
                break
        else:
            print(f"Error: No sheet found with required columns in {file_path}")
            return pd.DataFrame(), pd.DataFrame(), None

        # Extract hotel name
        hotel_name = df['hotel_name'].iloc[0] if 'hotel_name' in df.columns else None

        # Convert checkin_date to datetime and then to date (removing time component)
        df['checkin_date'] = pd.to_datetime(df['checkin_date']).dt.date

        # Create a date range for all days in the file
        min_date = df['checkin_date'].min()
        max_date = df['checkin_date'].max()
        all_dates = pd.date_range(start=min_date, end=max_date).date
        
        # Create DataFrames with all dates
        result_df = pd.DataFrame({'checkin_date': all_dates})
        detailed_df = pd.DataFrame({'checkin_date': all_dates})

        # Process each date
        for date in all_dates:
            price, room_name, breakfast, refundable, occupancy = process_date_with_criteria_hierarchy(df, date)
            result_df.loc[result_df['checkin_date'] == date, 'price'] = price
            detailed_df.loc[detailed_df['checkin_date'] == date, 'price'] = price
            detailed_df.loc[detailed_df['checkin_date'] == date, 'room_name'] = room_name
            detailed_df.loc[detailed_df['checkin_date'] == date, 'occupancy'] = occupancy
            detailed_df.loc[detailed_df['checkin_date'] == date, 'breakfast_included'] = breakfast
            detailed_df.loc[detailed_df['checkin_date'] == date, 'refundable'] = refundable

        return result_df, detailed_df, hotel_name
    except Exception as e:
        print(f"Error processing {file_path}: {str(e)}")
        return pd.DataFrame(), pd.DataFrame(), None

def process_date_with_criteria_hierarchy(df, date):
    occupancy_order = [2, 3, 4, 5, 1]
    
    for occupancy in occupancy_order:
        filtered_df = df[(df['checkin_date'] == date) & 
                         (df['type'] == 'Regular') & 
                         (df['occupancy'] == occupancy)]
        
        if not filtered_df.empty:
            # Find the lowest price
            min_price_row = filtered_df.loc[filtered_df['price'].idxmin()]
            return (remove_apostrophe(min_price_row['price']), 
                    min_price_row['name'], 
                    min_price_row['breakfast_included'], 
                    min_price_row['refundable'],
                    min_price_row['occupancy'])
    
    return 'Sold Out', 'N/A', 'N/A', 'N/A', 'N/A'

def main():
    directory = "./THKHA/THKHA-codes-06/Bhandari/"   # Replace with the actual directory path

    print(f"Looking for files in directory: {directory}")

    # Get all Excel files in the directory
    files = [f for f in os.listdir(directory) if f.endswith(".xlsx")]
    print(f"\nFound {len(files)} Excel files:")
    for file in files:
        print(f"  - {file}")

    if not files:
        print("\nNo Excel files found in the directory.")
        return

    # Process all files
    all_results = pd.DataFrame()
    all_detailed_results = pd.DataFrame()
    hotel_name = None
    for file in files:
        file_path = os.path.join(directory, file)
        df, detailed_df, file_hotel_name = process_hotel_file(file_path)
        if not df.empty:
            all_results = pd.concat([all_results, df], ignore_index=True)
            all_detailed_results = pd.concat([all_detailed_results, detailed_df], ignore_index=True)
        if hotel_name is None and file_hotel_name is not None:
            hotel_name = file_hotel_name

    if all_results.empty:
        print("No data found matching any criteria for any of the files.")
        return

    if hotel_name is None:
        print("Warning: Could not determine hotel name from files. Using 'Unknown Hotel'.")
        hotel_name = "Unknown Hotel"

    # Process results for both DataFrames
    for df in [all_results, all_detailed_results]:
        # Convert 'Sold Out' to a high number for sorting purposes
        df['sort_price'] = pd.to_numeric(df['price'], errors='coerce').fillna(float('inf'))

        # Sort the results by date and price
        df.sort_values(['checkin_date', 'sort_price'], inplace=True)

        # Remove any duplicate dates, keeping the cheapest price
        df.drop_duplicates(subset=['checkin_date'], keep='first', inplace=True)

        # Drop the temporary 'sort_price' column
        df.drop(columns=['sort_price'], inplace=True)

    # Rename columns for the simple report
    all_results.columns = ['Date', hotel_name]

    # Rename and reorder columns for the detailed report
    all_detailed_results.columns = ['Date', 'Price', 'Room Name', 'Occupancy', 'Breakfast Included', 'Refundable']
    all_detailed_results = all_detailed_results[['Date', 'Price', 'Room Name', 'Occupancy', 'Breakfast Included', 'Refundable']]

    # Convert Date column to string format 'YYYY-MM-DD' for both DataFrames
    all_results['Date'] = all_results['Date'].astype(str)
    all_detailed_results['Date'] = all_detailed_results['Date'].astype(str)

    # Save the results to new Excel files
    current_date = datetime.now().strftime("%Y%m%d")
    simple_output_file_path = f'{hotel_name}_prices_{current_date}.xlsx'
    detailed_output_file_path = f'{hotel_name}_detailed_prices_{current_date}.xlsx'

    try:
        all_results.to_excel(simple_output_file_path, index=False, engine='openpyxl')
        print(f"Simple results saved to {simple_output_file_path}")
        
        all_detailed_results.to_excel(detailed_output_file_path, index=False, engine='openpyxl')
        print(f"Detailed results saved to {detailed_output_file_path}")
    except Exception as e:
        print(f"Error saving results: {str(e)}")

if __name__ == "__main__":
    main()