import os
import sys
import pandas as pd

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
sys.path.insert(1, ROOT_DIR)
for path, dirs, files in os.walk(ROOT_DIR):
    if '__pycache__' in path:
        continue
    sys.path.insert(1, path)

from etc.config import path_output, file_52w, file_insight

# 52 Week Insights
import xlsxwriter
import xlrd
import openpyxl
from datetime import datetime as dt
import calendar

file_masterdata = os.path.join(path_output, file_insight)
file_52w = os.path.join(path_output, file_52w)
df = pd.read_excel(file_masterdata, sheet_name='Nifty-All')
output_file = file_52w
if os.path.exists(output_file):
    os.remove(output_file)

# Main Call
writer = pd.ExcelWriter(output_file, engine='xlsxwriter',
                        datetime_format='mmm dd yyyy',
                        date_format='mmmm dd yyyy',
                        options={'strings_to_urls': False, 'strings_to_formulas': False})

print(df.columns)
month_sell = df["52wHighDate"].dt.month.apply(lambda x: calendar.month_name[x]).value_counts().to_dict()
month_buy = df["52wLowDate"].dt.month.apply(lambda x: calendar.month_name[x]).value_counts().to_dict()

day_sell = df["52wHighDate"].dt.day.value_counts().to_dict()
day_buy = df["52wLowDate"].dt.day.value_counts().to_dict()

year_sell = df["52wHighDate"].dt.year.value_counts().to_dict()
year_buy = df["52wLowDate"].dt.year.value_counts().to_dict()

# right days to buy
week_day_sell = df["52wHighDate"].dt.weekday_name.value_counts().to_dict()
week_day_buy = df["52wLowDate"].dt.weekday_name.value_counts().to_dict()

quarter_sell = df["52wHighDate"].dt.quarter.value_counts().to_dict()
quarter_buy = df["52wLowDate"].dt.quarter.value_counts().to_dict()

# Long term investment week number
week_sell = df["52wHighDate"].dt.week.value_counts().to_dict()
week_buy = df["52wLowDate"].dt.week.value_counts().to_dict()

today = dt.now()
quarter = int(today.strftime('%m')) // 3 + 1
month = today.strftime('%B')
day = int(today.strftime('%d'))
week_num = int(today.strftime('%V'))
weekday = today.strftime('%A')

workbook = xlsxwriter.Workbook(output_file)
worksheet = workbook.add_worksheet('52W Stats')

# Set tab colors
worksheet.set_tab_color('#FF9900')  # Orange

workbook.set_properties({
    'title': '52Week Stats',
    'subject': 'With document properties',
    'author': 'John McNamara',
    'manager': 'Dr. Vijay',
    'company': 'Awesome Work Pays',
    'category': 'Example spreadsheets',
    'keywords': 'Sample, Example, Properties',
    'comments': 'Created with Python and XlsxWriter',
    'status': 'Quo',
})

diag_format = workbook.add_format({'diag_type': 3, 'diag_border': 1, 'diag_color': 'red', })
buy_format = workbook.add_format({'bold': True, 'font_color': 'green', 'align': 'center'})
sell_format = workbook.add_format({'bold': True, 'font_color': 'red', 'align': 'center'})
bold_format = workbook.add_format({'bold': True})
title_format = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'yellow', 'border': True})
sell1_format = workbook.add_format({'bg_color': '#FF5733', 'border': True})
sell2_format = workbook.add_format({'bg_color': '#F1948A'})
sell3_format = workbook.add_format({'bg_color': '#FADBD8'})
buy1_format = workbook.add_format({'bg_color': '#10CF07', 'border': True})
buy2_format = workbook.add_format({'bg_color': '#2ECC71'})
buy3_format = workbook.add_format({'bg_color': '#ABEBC6'})

worksheet.set_column('A:A')
worksheet.write('A1', "Month", title_format)
# worksheet.add_table('A2:D14', {'header_row': False, 'style': 'Table Style Light 11'})
worksheet.write('A2', "Buy", buy_format)
worksheet.write('C2', "Sell", sell_format)
row = 2
col = 0
# MONTH BUY
for item, cost in (list(month_buy.items())):
    if item == month:
        worksheet.write(row, col, item, buy1_format)
    else:
        worksheet.write(row, col, item, bold_format)
    worksheet.write(row, col + 1, cost)
    row += 1
row = 2
col = 2
# MONTH SELL
for item, cost in (list(month_sell.items())):
    if item == month:
        worksheet.write(row, col, item, buy1_format)
    else:
        worksheet.write(row, col, item, bold_format)
    worksheet.write(row, col + 1, cost)
    row += 1

worksheet.write('A16', "Weekday", title_format)
# worksheet.add_table('A17:D22', {'header_row': False, 'style': 'Table Style Light 11'})
worksheet.write('A17', "Buy", buy_format)
worksheet.write('C17', "Sell", sell_format)
row = 17
col = 0
# WEEKDAY BUY
for item, cost in (list(week_day_buy.items())):
    if item == weekday:
        worksheet.write(row, col, item, buy1_format)
    else:
        worksheet.write(row, col, item, bold_format)
    worksheet.write(row, col + 1, cost)
    row += 1
row = 17
col = 2
# WEEKDAY SELL
for item, cost in (list(week_day_sell.items())):
    if item == weekday:
        worksheet.write(row, col, item, sell1_format)
    else:
        worksheet.write(row, col, item, bold_format)
    worksheet.write(row, col + 1, cost)
    row += 1

worksheet.write('A24', "Quarter", title_format)
worksheet.add_table('A25:D29', {'header_row': False, 'style': 'Table Style Light 11'})
worksheet.write('A25', "Buy", buy_format)
worksheet.write('C25', "Sell", sell_format)
row = 25
col = 0
# Quarter BUY
for item, cost in (list(quarter_buy.items())):
    if item == quarter:
        worksheet.write(row, col, item, buy1_format)
    else:
        worksheet.write(row, col, item, bold_format)
    worksheet.write(row, col + 1, cost)
    row += 1
row = 25
col = 2
# Quarter SELL
for item, cost in (list(quarter_sell.items())):
    if item == quarter:
        worksheet.write(row, col, item, sell1_format)
    else:
        worksheet.write(row, col, item, bold_format)
    worksheet.write(row, col + 1, cost)
    row += 1

worksheet.write('F1', "Day", title_format)
# worksheet.add_table('F3:I33', {'header_row': False, 'style': 'Table Style Light 9'})
worksheet.write('F2', "Buy", buy_format)
worksheet.write('H2', "Sell", sell_format)
row = 2
col = 5
# Day BUY
for item, cost in (list(day_buy.items())):
    if int(item) == day:
        worksheet.write(row, col, item, buy1_format)
    elif int(item) == day + 1:
        worksheet.write(row, col, item, buy2_format)
    elif int(item) == day - 1:
        worksheet.write(row, col, item, buy3_format)
    else:
        worksheet.write(row, col, item, bold_format)
    worksheet.write(row, col + 1, cost)
    row += 1
row = 2
col = 7
# Day SELL
for item, cost in (list(day_sell.items())):
    if int(item) == day:
        worksheet.write(row, col, item, sell1_format)
    elif int(item) == day + 1:
        worksheet.write(row, col, item, sell2_format)
    elif int(item) == day - 1:
        worksheet.write(row, col, item, sell3_format)
    else:
        worksheet.write(row, col, item, bold_format)
    worksheet.write(row, col + 1, cost)
    row += 1

worksheet.write('K1', "Week", title_format)
# worksheet.add_table('K3:N50', {'header_row': False, 'style': 'Table Style Light 13'})
worksheet.write('K2', "Buy", buy_format)
worksheet.write('M2', "Sell", sell_format)
row = 2
col = 10
# Week BUY
for item, cost in (list(week_buy.items())):
    if int(item) == week_num:
        worksheet.write(row, col, item, buy1_format)
    elif int(item) == week_num + 1:
        worksheet.write(row, col, item, buy2_format)
    elif int(item) == week_num - 1:
        worksheet.write(row, col, item, buy3_format)
    else:
        worksheet.write(row, col, item, bold_format)
    worksheet.write(row, col + 1, cost)
    row += 1
row = 2
col = 12
# Week SELL
for item, cost in (list(week_sell.items())):
    if int(item) == week_num:
        worksheet.write(row, col, item, sell1_format)
    elif int(item) == week_num + 1:
        worksheet.write(row, col, item, sell2_format)
    elif int(item) == week_num - 1:
        worksheet.write(row, col, item, sell3_format)
    else:
        worksheet.write(row, col, item, bold_format)
    worksheet.write(row, col + 1, cost)
    row += 1

for buy in ['B3:B14', 'B18:B22', 'B26:B29', 'G3:G33', 'L3:L50']:
    worksheet.conditional_format(buy, {'type': 'data_bar', 'bar_color': '#63C384'})
for sell in ['D3:D14', 'D18:D22', 'D26:D29', 'I3:I33', 'N3:N50']:
    worksheet.conditional_format(sell, {'type': 'data_bar', 'bar_color': '#FF9999'})

writer = pd.ExcelWriter(output_file, engine='openpyxl')
if os.path.exists(output_file):
    book = openpyxl.load_workbook(output_file)
    writer.book = book

df.to_excel(writer, sheet_name='Nifty', index=False)

writer.save()
writer.close()
print('Written')