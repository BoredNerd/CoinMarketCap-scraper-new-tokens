from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.chart import PieChart, Reference
import subprocess
import os
import datetime
import requests
import pathlib

html_text = requests.get("https://coinmarketcap.com/new/").text
soup = BeautifulSoup(html_text, "lxml")

wb = Workbook()
ws = wb.active
ws.append(["Name", "Price", "MarketCap"])

token = soup.find_all("tr")
for details in token:
    name_token = details.find("p", class_="sc-1eb5slv-0 iworPT")
    price_token = details.find_all("span")
    market_cap = details.find_all("td")
    if name_token and price_token is not None:
        print(f"Name : {name_token.text}")
        print(f"Price : {price_token[2].text}")
        price_market = market_cap[6].text
        print(f"Market cap : {price_market}")
        if price_market != "--":
            ws.append([name_token.text, price_token[2].text, int(price_market.replace(",", "").replace("$", ""))])
        print(" ")
ws.column_dimensions['A'].width = 30
ws.column_dimensions['B'].width = 30
ws.column_dimensions['C'].width = 30

for col in ws['C']:
   if col != "MarketCap":
       ws[col.coordinate].number_format = '#,##0.00$'
chart = PieChart()
categories = Reference(ws, min_col = 1, min_row=2, max_row= ws.max_row)
data = Reference(ws, min_col = 3, min_row=2, max_row= ws.max_row)

chart.add_data(data)
chart.set_categories(categories)
chart.title = "New Tokens"
chart.height = 20
chart.width = 30
ws.add_chart(chart, "E2")

if not os.path.exists("token_history"):
    os.makedirs("token_history")
namefile = "token_history/NewToken_" + datetime.datetime.now().strftime("%b-%d-%Y_%H-%M") + ".xlsx"
wb.save(namefile)
openExcel = (os.path.abspath(os.getcwd()) + "\\" + namefile).replace("\\","\\\\").replace("/","\\\\")

subprocess.Popen(openExcel , shell="false")


