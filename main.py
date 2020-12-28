from openpyxl import Workbook, styles
import os
import requests

#--- Creating sheet and styles ---#

wb = Workbook()
ws = wb.active
ws.title = "Summary"
ws['A1'] = 'Summary Page'
title_style = styles.Font(bold=True, sz=30, underline='single', italic=True)
ws['A1'].font = title_style
wb.save('stock_tracker.xlsx')

#TODO - Get stock prices
#TODO - Create a Google sheet version

#------- STOCK API & VARIABLES (TO HASH LATER) ---------#

STOCK = "AAPL"
COMPANY_NAME = "Apple"

stock_api_key = 'P80S76BFMQ83FJBH'

STOCK_ENDPOINT = "https://www.alphavantage.co/query"

stock_param = {
    "function": "TIME_SERIES_DAILY_ADJUSTED",
    "symbol": STOCK,
    "apikey": stock_api_key,
}

#------- STOCK API CALLS ---------#

response = requests.get(STOCK_ENDPOINT, params=stock_param)
response.raise_for_status()
data = response.json()["Time Series (Daily)"]

today = list(data.keys())[0]
yesterday = list(data.keys())[1]

#------- RUN STOCK REQUESTS ---------#

def value_check():
    today_open = float((data[today]['1. open']))
    today_close = float((data[yesterday]['1. open']))
    change = round(((today_open - today_close) / today_open)*100,2)
    print(f"Today's stock price for {COMPANY_NAME} : {today_open}")
    print(f"Yesterday's price for {COMPANY_NAME}: {today_close}")
    print(f"Percentage change is {change}%")

value_check()
