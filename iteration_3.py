import requests
import pandas as pd
from datetime import datetime
import pytz
import os
import shutil
import configparser
import xlsxwriter

# Function to read stock symbols from the text file
def read_stock_symbols(file_path):
    try:
        with open(file_path, 'r') as file:
            stock_symbols = file.read().strip().split('\n')
            return [symbol.strip() for symbol in stock_symbols]
    except Exception as e:
        print(f"Error reading the file: {e}")
        return []

# Function to fetch stock data using Finnhub API
def fetch_stock_data(symbol, api_key, exchange_rate):
    base_url = "https://finnhub.io/api/v1/quote"
    try:
        response = requests.get(f"{base_url}?symbol={symbol}&token={api_key}")
        if response.status_code == 200:
            data = response.json()
            open_price_usd = data.get('o', 0)
            current_price_usd = data.get('c', 0)
            open_price_eur = round(open_price_usd * exchange_rate, 2)
            current_price_eur = round(current_price_usd * exchange_rate, 2)
            rise_percentage = round(((current_price_eur - open_price_eur) / open_price_eur * 100), 1) if open_price_eur else 0
            return {
                "symbol": symbol,
                "stock_name": symbol,
                "open_price": open_price_eur,
                "current_price": current_price_eur,
                "rise%": rise_percentage
            }
        else:
            print(f"Failed to fetch data for {symbol}: HTTP {response.status_code}")
            return None
    except Exception as e:
        print(f"Error fetching data for {symbol}: {e}")
        return None

# Function to fetch USD to EUR exchange rate from Finnhub API
def fetch_exchange_rate(api_key, base_currency="USD", target_currency="EUR"):
    """
    Fetches the exchange rate between the base currency and the target currency.
    Args:
        api_key (str): Your API key for ExchangeRate-API.
        base_currency (str): The currency to convert from (default: USD).
        target_currency (str): The currency to convert to (default: EUR).

    Returns:
        float: The exchange rate if successful, or None if an error occurs.
    """
    url = f"https://v6.exchangerate-api.com/v6/{api_key}/latest/{base_currency}"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            conversion_rates = data.get("conversion_rates", {})
            return conversion_rates.get(target_currency)
        else:
            print(f"Error: HTTP {response.status_code}, {response.text}")
            return None
    except Exception as e:
        print(f"Error fetching exchange rate: {e}")
        return None

# API key and file paths
api_key = "ct2rhrpr01qiurr42g40ct2rhrpr01qiurr42g4g"
current_date = datetime.now().strftime('%Y-%m-%d')
previous_file = f'stock_data_output_{current_date}.xlsx'
output_file = f'stock_data_output_{current_date}.xlsx'

def get_rise_threshold():
    config = configparser.ConfigParser()
    config.read('config.ini')
    try:
        rise_threshold = float(config['settings']['rise_threshold'])
    except KeyError:
        print("Key 'rise_threshold' not found in config.ini. Using default value of 2.")
        rise_threshold = 2
    except ValueError:
        print("Invalid 'rise_threshold' value in config.ini. Using default value of 2.")
        rise_threshold = 2
    return rise_threshold

rise_threshold = get_rise_threshold()

# Fetch the exchange rate
exchange_rate_api_key = "a42ed75da2b0c4d17ba4e4c0"  # Replace with your actual API key
exchange_rate = fetch_exchange_rate(exchange_rate_api_key)

# Read stock symbols
stock_symbols = read_stock_symbols('new_stocks.txt')
stock_data = [fetch_stock_data(symbol, api_key, exchange_rate) for symbol in stock_symbols if fetch_stock_data(symbol, api_key, exchange_rate)]
stocks_df = pd.DataFrame(stock_data)
stocks_df["Execution"] = "Current"

# Add CET timestamp
cet_timezone = pytz.timezone('CET')
last_refreshed_time = datetime.now(cet_timezone).strftime('%Y-%m-%d %H:%M:%S CET')

# Handle previous file and `NewStock` sheet
if os.path.exists(previous_file):
    with pd.ExcelFile(previous_file) as xls:
        if 'NewStock' in xls.sheet_names:
            previous_stocks_df = pd.read_excel(xls, sheet_name='NewStock')
        else:
            previous_stocks_df = pd.DataFrame(columns=stocks_df.columns)
        if 'WatchedStock' in xls.sheet_names:
            previous_watched_df = pd.read_excel(xls, sheet_name='WatchedStock')
        else:
            previous_watched_df = pd.DataFrame(columns=stocks_df.columns)

        previous_stocks_df = pd.concat([previous_stocks_df, previous_watched_df]).reset_index(drop=True)
        previous_stocks_df['Execution'] = "Previous"

else:
    previous_stocks_df = pd.DataFrame(columns=stocks_df.columns)

# Clean and prepare data for merging
previous_stocks_df["rise%"] = previous_stocks_df["rise%"].round(2).astype(str)
stocks_df["rise%"] = stocks_df["rise%"].round(2).astype(str)
previous_stocks_df["rise%"] = previous_stocks_df["rise%"].str.rstrip('%').astype(float)
stocks_df["rise%"] = stocks_df["rise%"].str.rstrip('%').astype(float)

# Merge current and previous stock data on symbol to calculate the difference
merged_df = stocks_df.merge(
    previous_stocks_df, on="symbol", suffixes=("_curr", "_prev"), how="left"
)

# Calculate the diff% and update the stocks_df with the diff column
merged_df["curr_ct_price - prev_ct_price"] = (merged_df["current_price_curr"] - merged_df["current_price_prev"]).round(1)

# Create the final dataframe with diff% and appropriate columns
final_df = merged_df[[ "symbol", "stock_name_curr", "open_price_curr", "current_price_curr", "rise%_curr", "Execution_curr", "curr_ct_price - prev_ct_price"
]].rename(columns={
    "stock_name_curr": "stock_name",
    "open_price_curr": "open_price",
    "current_price_curr": "current_price",
    "rise%_curr": "rise%",
    "Execution_curr": "Execution"
})

# Concatenate previous and current stock data, drop duplicates, and reset index
combined_stocks_df = pd.concat([previous_stocks_df, final_df]).drop_duplicates(subset="symbol", keep="last")
combined_stocks_df.reset_index(drop=True, inplace=True)

# Split the combined_stocks_df based on the rise% condition
new_stock_df = combined_stocks_df[combined_stocks_df['rise%'] > rise_threshold].copy()
watched_stock_df = combined_stocks_df[combined_stocks_df['rise%'] <= rise_threshold].copy()

#watched_stock_df = watched_stock_df.drop(columns=['Execution'], errors='ignore')

# Define a function to apply color formatting to the 'Execution' column
def color_execution(row):
    color = 'color: green' if row['Execution'] == 'Current' else 'color: orange' if row['Execution'] == 'Previous' else ''
    return [color] * len(row)  # Apply the color to all columns in the row

# Apply the color formatting to the entire row based on the 'Execution' column
new_stock_df_style = new_stock_df.style.apply(color_execution, axis=1)
watched_stock_df_style = watched_stock_df.style.apply(color_execution, axis=1)

# Read my_stock symbols and data
my_stock_symbols = read_stock_symbols('my_stock.txt')
my_stock_data = [fetch_stock_data(symbol, api_key, exchange_rate) for symbol in my_stock_symbols if fetch_stock_data(symbol, api_key, exchange_rate)]
my_stock_df = pd.DataFrame(my_stock_data)

# Write to Excel with color formatting
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Apply the styles before writing to Excel
    new_stock_df_style.to_excel(writer, index=False, sheet_name='NewStock')
    watched_stock_df_style.to_excel(writer, index=False, sheet_name='WatchedStock')
    my_stock_df.to_excel(writer, index=False, sheet_name='MyStock')
    workbook = writer.book
    tz_worksheet = workbook.add_worksheet('Timezonesheet')
    tz_worksheet.write(0, 0, 'Last Refreshed Time')
    tz_worksheet.write(1, 0, last_refreshed_time)
    tz_worksheet.write(0, 1, 'CET timezone')
    tz_worksheet.write(1, 1, 'CET')

print(f"Data written to {output_file} successfully.")
# Define the final file name to copy to
final_output_file = 'stock_data_output_1.xlsx'

# Copy the output file to stock_data_output_1.xlsx
try:
    shutil.copyfile(output_file, final_output_file)
    print(f"File copied to {final_output_file} successfully.")
except Exception as e:
    print(f"Error copying the file: {e}")