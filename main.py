from openpyxl import Workbook, styles
from openpyxl.chart import LineChart,Reference
from openpyxl.chart.axis import DateAxis
import requests

###------- CREATING SHEET AND STYLES ---------###

wb = Workbook()
ws = wb.active
ws.title = "Summary"
ws['A1'] = 'Summary Page'
title_style = styles.Font(bold=True, sz=30, underline='single', italic=True)
ws['A1'].font = title_style
apple_tab = wb.create_sheet("Apple Tab")
apple_tab['A1'] = "Apple raw data"
apple_tab['A1'].font = title_style


###------- API VARIABLES ---------###

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
stock_data = response.json()["Time Series (Daily)"]

today = list(stock_data.keys())[0]
yesterday = list(stock_data.keys())[1]

#------- RUN STOCK REQUESTS ---------#

apple_date = []
apple_low = []
apple_high = []

#TODO: Fix bug below to get data into a format ready for excel chart.

def value_check():
    for i in stock_data.keys():
        high_price = float((stock_data[i]['2. high']))
        low_price = float((stock_data[i]['3. low']))
        apple_date.append(i)
        apple_low.append(low_price)
        apple_high.append(high_price)

value_check()

print(apple_daily_prices)


#------- CREATING CHART ---------#

rows = [
    ['Date', 'Low', 'High']]


for row in rows:
    ws.append(row)
    for i in range(len(apple_daily_prices)):
        ws.append(apple_daily_prices[i])


c1 = LineChart()
c1.title = "Test Chart"
c1.style = 13
c1.y_axis.title = 'Size'
c1.x_axis.title = 'Test Number'

data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=7)
c1.add_data(data, titles_from_data=True)

# Style the lines
s1 = c1.series[0]
s1.marker.symbol = "triangle"
s1.marker.graphicalProperties.solidFill = "FF0000" # Marker filling
s1.marker.graphicalProperties.line.solidFill = "FF0000" # Marker outline

s1.graphicalProperties.line.noFill = True

s2 = c1.series[1]
s2.graphicalProperties.line.solidFill = "00AAAA"
s2.graphicalProperties.line.dashStyle = "sysDot"
s2.graphicalProperties.line.width = 100050 # width in EMUs

s2 = c1.series[2]
s2.smooth = True # Make the line smooth

ws.add_chart(c1, "E10")

wb.save('stock_tracker.xlsx')