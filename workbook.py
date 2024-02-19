import yfinance as yf
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from datetime import datetime,timedelta

# Get the current date and time
current_datetime = datetime.now()

# Extract the date component
current_date = current_datetime.date()- timedelta(days=1)

# Subtract five years from the current date
five_years_ago = current_datetime - timedelta(days=365 * 5)

# Extract the date component
five_years_ago_date = five_years_ago.date()

def get_historical_data(ticker, start_date, end_date):
    stock_data = yf.download(ticker, start=start_date, end=end_date)
    return stock_data

def create_workbook(tickers, start_date, end_date):
    wb = Workbook()
    for ticker in tickers:
        stock_data = get_historical_data(ticker, start_date, end_date)
        ws = wb.create_sheet(title=ticker)
        df = pd.DataFrame(stock_data)
        for r in dataframe_to_rows(df, index=True, header=True):
            ws.append(r)
    return wb

if __name__ == "__main__":
    tickers = ['BALRAMCHIN.NS', 'RENUKA.NS', 'AVADHSUGAR.NS','BAJAJHIND.NS','UTTAMSUGAR.NS','TRIVENI.NS']
    start_date = five_years_ago_date
    end_date = current_date
    wb = create_workbook(tickers, start_date, end_date)
    wb.save('stock_historical_data.xlsx')
