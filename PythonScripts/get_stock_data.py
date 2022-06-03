from pathlib import Path  # Python Standard Library

import xlwings as xw  # pip install xlwings
import yfinance as yf  # pip install yfinance

excel_file = Path(__file__).parents[1] / "RunPython_Example.xlsx"

# Connect to workbook
wb = xw.Book(excel_file)
sht = wb.sheets["StockData"]
ticker = sht["TICKER"].value
period = sht["PERIOD"].value
headers = sht["A4"]
output_cell = sht["A5"]

# Init yf ticket object
data = yf.Ticker(ticker)

print("Clear Excel cell range...")
output_cell.expand().clear_contents()

print("Inserting headers...")
headers.value = f"{period} market data for {ticker}"

print(f"Get {period} historical market data for: {ticker}")
hist = data.history(period=period)

print("Export historical market data to Excel...")
output_cell.value = hist
