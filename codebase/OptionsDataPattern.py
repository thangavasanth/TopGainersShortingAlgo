# -*- coding: utf-8 -*-
"""
Created on Thu May 22 21:52:15 2025

@author: thang
"""

import xlwings as xw
import pandas as pd
import math
import datetime
from ib_insync import IB, Stock, util
from collections import defaultdict


import xlwings as xw
from ib_insync import *
import pandas as pd
import os
from datetime import datetime, timedelta, time
import pytz
import subprocess
import math
import sys
import tempfile
from Crypto.Cipher import AES
util.startLoop()
import pandas as pd
from ib_insync import IB, Stock, util
import datetime
import os

# === Initialize IB connection ===
ib = IB()
ib.connect('127.0.0.1', 4001, 1234)
util.startLoop()

# === Excel setup ===
excel_path = r'{Yaali_Algo_Trading_Full_FilePath}'
wb = xw.Book(excel_path)
ws_options = wb.sheets['Options']

# === Read tickers from column B starting row 9 ===
tickers = ws_options.range('B9').expand('down').value
tickers = [t for t in tickers if t]

weekly_data = []
dropdowns = defaultdict(lambda: {"drop": [], "up": []})
current_prices = {}

# === Generate weekly data for each ticker ===
for ticker in tickers:
    try:
        print(f"Fetching data for {ticker}...")
        stock = Stock(ticker, 'SMART', 'USD')
        end_date = datetime.datetime.now()

        bars = ib.reqHistoricalData(
            stock,
            endDateTime=end_date,
            durationStr='2 Y',
            barSizeSetting='1 week',
            whatToShow='MIDPOINT',
            useRTH=True
        )

        df = util.df(bars)
        df['date'] = pd.to_datetime(df['date'])
        df.set_index('date', inplace=True)

        # ✅ Use open-to-close for accurate weekly price movement
        df['weekly_pct_change'] = ((df['close'] - df['open']) / df['open']) * 100
        
        # ✅ Capture current price for summary
        current_price = round(df['close'].iloc[-1], 2)
        current_prices[ticker] = current_price
        
        # Add metadata columns
        df['Ticker'] = ticker
        df['Week Period'] = df.index.strftime('%Y-%m-%d') + " to " + df.index.strftime('%Y-%m-%d')
        df['Week No'] = [(d - df.index[0]).days // 7 + 1 for d in df.index]
        df['Current Price'] = current_price
        
        # ✅ Truncate % change to keep only integer part toward zero
        # Rounded change
        # --- Round values ---
        df['Rounded Change'] = df['weekly_pct_change'].apply(lambda x: int(x) if pd.notnull(x) else pd.NA)
        
        # === DROP % (Plunge) ===
        drop_series = df['Rounded Change'][df['Rounded Change'] < 0].dropna()
        drop_exact = drop_series.value_counts().sort_index(ascending=True)  # small to big
        drop_cumulative = drop_exact.cumsum()
        
        drop_values = []
        default_drop = ""
        for val in drop_exact.index:
            exact = drop_exact[val]
            cumulative = drop_cumulative[val]  # correctly cumulative ≤ val
            formatted = f"'{val} ({exact})({cumulative})"  # <-- prefix with '
            drop_values.append(formatted)
            if not default_drop and cumulative >= 5:
                default_drop = formatted
        
        dropdowns[ticker]["drop"] = drop_values
        dropdowns[ticker]["default_drop"] = default_drop
        
        # === UP % (Surge) ===
        up_series = df['Rounded Change'][df['Rounded Change'] > 0].dropna()
        up_exact = up_series.value_counts().sort_index(ascending=False)  # big to small
        up_cumulative = up_exact.cumsum()
        
        up_values = []
        default_up = ""
        for val in up_exact.index:
            exact = up_exact[val]
            cumulative = up_cumulative[val]  # cumulative ≥ val
            formatted = f"{val} ({exact})({cumulative})"
            up_values.append(formatted)
            if not default_up and cumulative >= 5:
                default_up = formatted
        
        dropdowns[ticker]["up"] = up_values
        dropdowns[ticker]["default_up"] = default_up




        # Add rows to output table
        for idx, row in df.iterrows():
            weekly_data.append({
                "Ticker": ticker,
                "Week Period": row['Week Period'],
                "Week No": row['Week No'],
                "Weekly Open Price": row['open'],
                "Weekly End Price on Friday EOD": row['close'],
                "Percentage Change": row['weekly_pct_change'],
                "Current Price": current_price
            })

    except Exception as e:
        print(f"Error processing {ticker}: {e}")
        continue

# === Write to "Options Data" sheet ===
df_weekly = pd.DataFrame(weekly_data)
if 'OptionsData' not in [s.name for s in wb.sheets]:
    ws_data = wb.sheets.add('OptionsData')
else:
    ws_data = wb.sheets['OptionsData']
ws_data.clear()
ws_data.range('A1').value = [df_weekly.columns.tolist()] + df_weekly.values.tolist()

# === Apply dropdowns and formulas to "Options" sheet ===
drop_col, up_col = 'C', 'D'
put_buf_col, call_buf_col = 'E', 'F'
curr_price_col, put_strike_col, call_strike_col = 'G', 'H', 'I'

for i, ticker in enumerate(tickers, start=9):
    if not ticker:
        break  # ✅ skip empty rows
    
    # Drop % dropdown and default value
    drop_values = ','.join(dropdowns[ticker]["drop"])
    drop_cell = ws_options.range(f'{drop_col}{i}')
    if drop_values:
        ws_options.range(f'{drop_col}{i}').api.Validation.Delete()
        ws_options.range(f'{drop_col}{i}').api.Validation.Add(3, 1, 1, f'"{drop_values}"')
        if dropdowns[ticker]['default_drop']:
            ws_options.range(f'{drop_col}{i}').value = dropdowns[ticker]['default_drop']
            drop_cell.color = (248, 203, 173)  # RGB for #F8CBAD
    # Up % dropdown and default value
    up_values = ','.join(dropdowns[ticker]["up"])
    up_cell = ws_options.range(f'{up_col}{i}')
    if up_values:
        ws_options.range(f'{up_col}{i}').api.Validation.Delete()
        ws_options.range(f'{up_col}{i}').api.Validation.Add(3, 1, 1, f'"{up_values}"')
        if dropdowns[ticker]['default_up']:
            ws_options.range(f'{up_col}{i}').value = dropdowns[ticker]['default_up']
            up_cell.color = (198, 224, 180)  # #C6E0B4

    
    ws_options.range(f'{put_buf_col}{i}').value =1
    ws_options.range(f'{call_buf_col}{i}').value  =1 
    ws_options.range(f'{curr_price_col}{i}').value = current_prices[ticker]

    ws_options.range(f'{put_strike_col}{i}').formula = (f'=G{i}*(1 - (ABS(LEFT(C{i}, FIND("(", C{i})-2)-E{i})/100))')
    ws_options.range(f'{call_strike_col}{i}').formula = (f'=G{i}*(1 + (ABS(LEFT(D{i}, FIND("(", D{i})-2)+F{i})/100))')

        

# === Save and close ===
wb.save()
ib.disconnect()
# wb.app.quit()
sys.exit()