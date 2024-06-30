from src.tradingview_screener.screener import Query
from src.tradingview_screener.query import Column
from openpyxl import load_workbook
import os
import requests
from datetime import datetime, timedelta

alpha_vantage_key = "6XSD0XKGJCSLNBRR"
interval = '1min'
outputsize = 'full'
file_path = 'stock_data.xlsx'

n_rows, df = (Query()
 .select('name', 'close', 'premarket_close', 'premarket_change', 'premarket_volume')
 .where(
    Column('premarket_change') > 10,
    Column('premarket_volume') > 3_000_000,
    Column('premarket_close') > 0.9,
    Column('exchange').isin(['NASDAQ', 'NYSE'])
 ).get_scanner_data())

tickers = df["name"].tolist()

for ticker in tickers:
   today = datetime.today()
   latest_trading_day = today

   while latest_trading_day.weekday() >= 5:
      latest_trading_day -= timedelta(days=1)

   latest_trading_day_str = latest_trading_day.strftime('%Y-%m-%d')

   url = f'https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol={ticker}&interval={interval}&outputsize={outputsize}&apikey={alpha_vantage_key}'
   response = requests.get(url)

   if response.status_code == 200:
      data = response.json()

      # Extract time series data
      time_series = data[f'Time Series ({interval})']
      first_15_bars = []

      for timestamp, bar_data in time_series.items():
         if timestamp.startswith(latest_trading_day_str):
            timestamp_time = timestamp.split()[1]
            if '09:30:00' <= timestamp_time < '09:45:00':
               first_15_bars.append(bar_data)
               if len(first_15_bars) == 15:
                  break

      if len(first_15_bars) < 15:
         print(f"Warning: Less than 15 bars found for {latest_trading_day_str}")

      lastBar = first_15_bars[0]
      firstBar = first_15_bars[-1]
      
      openFirstBar = float(firstBar["1. open"])
      closeLastBar = float(lastBar["4. close"])
      
      percentageGain = ((closeLastBar - openFirstBar) / openFirstBar) * 100
      totalGain = (100 / 100) * (percentageGain + 100)
      
      wb = load_workbook(filename=file_path)
      sheet = wb.active
      
      new_data = [latest_trading_day_str, ticker, openFirstBar, closeLastBar, percentageGain, totalGain]
      sheet.append(new_data)
      wb.save(file_path)
      
   else:
      print(f"Error fetching data. Status code: {response.status_code}")


