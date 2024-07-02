from src.tradingview_screener.screener import Query
from src.tradingview_screener.query import Column
from openpyxl import load_workbook
from datetime import datetime, timedelta, timezone, time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import config
import smtplib, ssl
import requests
import time
import schedule

result_of_today = []

def run_daily_task():
   _, df = (Query()
   .select('name', 'close', 'premarket_close', 'premarket_change', 'premarket_volume')
   .where(
      Column('premarket_change') > 10,
      Column('premarket_volume') > 3_000_000,
      Column('premarket_close') > 0.9,
      Column('exchange').isin(['NASDAQ', 'NYSE'])
   ).get_scanner_data())

   # TO-DO
   # make a dictionary to store more premarket data.
   tickers = df["name"].tolist()

   for ticker in tickers:
      today = datetime.today()
      
      if datetime.now().time() > time(8, 0, 0):
         print("[Project Mercury] today's time has passed 6am.")
         return
      while today.weekday() >= 5:
         today -= timedelta(days=1)
         
      latest_trading_day_str = today.strftime('%Y-%m-%d')
      url = f'https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol={ticker}&interval={config.interval}&outputsize={config.outputsize}&apikey={config.alpha_vantage_key}'
      response = requests.get(url)

      if response.status_code == 200:
         data = response.json()
         time_series = data[f'Time Series ({config.interval})']
         first_15_bars = []

         for timestamp, bar_data in time_series.items():
            if timestamp.startswith(latest_trading_day_str):
               timestamp_time = timestamp.split()[1]
               if '09:30:00' <= timestamp_time < '09:45:00':
                  first_15_bars.append(bar_data)
                  if len(first_15_bars) == 15:
                     break

         if len(first_15_bars) < 15:
            print(f"[Project Mercury] Warning: Less than 15 bars found for {latest_trading_day_str}")

         lastBar = first_15_bars[0]
         firstBar = first_15_bars[-1]
         
         openFirstBar = float(firstBar["1. open"])
         closeLastBar = float(lastBar["4. close"])
         
         percentageGain = ((closeLastBar - openFirstBar) / openFirstBar) * 100
         totalGain = (100 / 100) * (percentageGain + 100)
         
         wb = load_workbook(filename=config.file_path)
         sheet = wb.active
         
         # TO-DO
         # can we also get the market cap and avg volume by any chance?
         new_data = [latest_trading_day_str, ticker, openFirstBar, closeLastBar, percentageGain, totalGain]
         sheet.append(new_data); result_of_today.append(new_data)
         wb.save(config.file_path)
      else:
         print(f"[Project Mercury] Error fetching data. Status code: {response.status_code}")
   
   try:    
      msg = MIMEMultipart()
      msg['From'] = f"{config.smtp_user}"
      msg['To'] = f"{config.to_email}"
      msg['Subject'] = "[Project Mercury] trading confirmation"
      msg.attach(MIMEText(f"{str(result_of_today)}", 'plain'))
      
      with smtplib.SMTP(config.smtp_server, port=config.smtp_port) as smtp:
         smtp.starttls()
         smtp.login(config.smtp_user, config.smtp_password)
         smtp.send_message(msg)
   except Exception as e:
      print(f"{e}")
   
def schedule_daily_task():
    schedule.every().day.at("07:30").do(run_daily_task)
    print("[Project Mercury] launching Project Mercury [executable at 7:30]!")
    
if __name__ == "__main__":
    schedule_daily_task()
    while True:
        schedule.run_pending()
        time.sleep(1)