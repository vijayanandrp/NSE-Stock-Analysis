# *-* coding: utf-8 *-*

"""
    Insights and Formatting
"""

import os
import sys
import pandas as pd

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
sys.path.insert(1, ROOT_DIR)
for path, dirs, files in os.walk(ROOT_DIR):
    if '__pycache__' in path:
        continue
    sys.path.insert(1, path)

from etc.config import path_output, file_insight, file_masterdata

columns_order = [
    'Series', 'Industry', 'Company Name', 'Symbol', '52wDiff', '52wHigh', '52wLow', '52wMid',
    'avgPrice', 'avg-52Low', 'avg-52Mid', 'CurRet%', 'MidRet%', 'LowRet%', 'dayDiff',
    'dayHigh', 'dayLow', 'change', 'pChange', 'open', 'lastPrice', 'prevClose',
    '52wHighDate', '52wLowDate', '52wDateDiff', 'ALGO_SCORE', 'TAG_SCORE', 'Nifty-Index',
    'totalBuyQuantity', 'totalSellQuantity',
    'totalTradedValue', 'totalTradedVolume', 'deliveryQuantity', 'deliveryToTradedQuantity', 'quantityTraded',
    'priceBand', 'pricebandlower', 'pricebandupper',
    'extremeLossMargin', 'faceValue', 'indexVar',
    'securityVar', 'surv_indicator', 'varMargin', 'adhocMargin',
    'applicableMargin', 'cm_ffm', 'isExDateFlag', 'ndEndDate', 'ndStartDate',
    'bcEndDate', 'bcStartDate', 'recordDate', 'secDate', 'exDate', 'purpose']

rename_columns = {'previousClose': 'prevClose',
                  'averagePrice': 'avgPrice',
                  'cm_adj_high_dt': '52wHighDate',
                  'cm_adj_low_dt': '52wLowDate',
                  'high52': '52wHigh',
                  'low52': '52wLow',
                  'TAGS': 'Nifty-Index'
                  }


def modify_columns(df):
    df.rename(index=str, columns=rename_columns, inplace=True)  # rename
    return df


def reorder_columns(df, extra_col=False):
    all_col = list(df.columns)
    diff_col = set(all_col) - set(columns_order)
    diff_col = list(diff_col)
    diff_col.sort()
    print(diff_col)
    if extra_col:
        df = df[columns_order + diff_col]
    else:
        df = df[columns_order]
    return df


def convert_float(x):
    try:
        return float(x)
    except:
        return float(0.0)


def parse_nifty(x):
    return x.replace('nifty', '')


def cost_per_range(cost):
    if 0 <= cost < 50:
        return 0
    elif 50 <= cost < 100:
        return 50
    elif 100 <= cost < 200:
        return 100
    elif 200 <= cost < 400:
        return 200
    elif 400 <= cost < 600:
        return 400
    elif 600 <= cost < 800:
        return 600
    elif 800 <= cost < 1000:
        return 800
    elif 1000 <= cost:
        return 1000


def format_nifty(writer_obj, sheet_name, df, close=False):
    num_rows = df.shape[0] + 1
    num_cols = df.shape[1] + 1
    print('Formating the columns - {}'.format(sheet_name))
    writer = writer_obj
    workbook = writer.book
    writer.book.strings_to_formulas = False
    cell_fmt = workbook.add_format({'align': 'bottom', 'bold': True, 'border': True})
    title_fmt = workbook.add_format({'bold': True, 'bg_color': '#59ddff', 'align': 'center',
                                     'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1,
                                     'font_color': 'black'})
    worksheet = writer.sheets[sheet_name]
    # Hide all rows without data.
    worksheet.set_default_row(hide_unused_rows=True)
    worksheet.freeze_panes(1, 0)
    worksheet.set_zoom(80)
    # green - high - red - low
    high_color_codes = {100: '#00cc00', 80: '#00ff00', 60: '#66ff99', 40: '#ccffcc', 20: '#e6ffe6', 0: '#ffffff'}
    low_color_codes = {100: '#ff0000', 80: '#ff794d', 60: '#ff4d4d', 40: '#ff8080', 20: '#ffcccc', 0: '#ffffff'}
    cost_color_codes = {1000: '#009933', 800: '#00b33c', 600: '#00e64d', 400: '#1aff66', 200: '#4dff88',
                        100: '#80ffaa', 50: '#b3ffcc', 0: '#ffffff'}
    # column
    column_width = {
        'A': 5, 'AA': 11, 'B': 29, 'AB': 66, 'C': 45, 'AC': 11, 'D': 14, 'AD': 11, 'E': 11, 'AE': 11, 'F': 11,
        'AF': 11, 'G': 11, 'AG': 11, 'H': 11, 'AH': 11, 'I': 11, 'AI': 11, 'J': 11, 'AJ': 11, 'K': 11, 'AK': 11,
        'L': 11, 'BE': 11,
        'AL': 11, 'M': 11, 'AM': 11, 'N': 11, 'AN': 11, 'O': 11, 'AO': 11, 'P': 11, 'AP': 11, 'Q': 11, 'AQ': 11,
        'R': 18, 'AR': 11, 'S': 18, 'AS': 11, 'T': 11, 'AT': 11, 'U': 12, 'AU': 11, 'V': 12, 'AV': 11, 'W': 11,
        'AW': 11, 'X': 11, 'AX': 11, 'Y': 13, 'AY': 11, 'Z': 13, 'AZ': 11, 'BA': 11, 'BB': 11, 'BC': 11, 'BD': 11,
        'BF': 11}
    # 'BG': 11, 'BH': 11, 'BI': 11, 'BJ': 11, 'BK': 11, 'BL': 11, 'BM': 11, 'BN': 11, 'BO': 11,
    # 'BP': 11, 'BQ': 11, 'BR': 11, 'BS': 11, 'BT': 11, 'BU': 11, 'BV': 11, 'BW': 11,
    # 'BX': 11, 'BY': 11, 'BZ': 8
    worksheet.set_row(0, 18, title_fmt)
    worksheet.autofilter('A1:BK{}'.format(num_rows))
    for idx, key in enumerate(column_width.keys()):
        worksheet.conditional_format("{}1".format(key), {'type': 'no_errors', 'format': title_fmt})
    for key in column_width.keys():
        worksheet.set_column("{}:{}".format(key, key), column_width[key], cell_fmt)
    # Per change
    worksheet.conditional_format('S2:S{}'.format(num_rows), {'type': 'data_bar', 'data_bar_2010': True,
                                                             'bar_negative_color_same': True,
                                                             'bar_negative_border_color_same': True})
    worksheet.conditional_format('R2:R{}'.format(num_rows), {'type': 'data_bar', 'data_bar_2010': True,
                                                             'bar_negative_color_same': True,
                                                             'bar_negative_border_color_same': True})
    # Day Diff Price
    worksheet.conditional_format('O2:O{}'.format(num_rows), {'type': 'data_bar'})
    # Days Diff Date, TAGS, ALGO
    for alpha in ['AA', 'Y', 'Z']:
        worksheet.conditional_format('{alpha}{start}:{alpha}{end}'.format(start=2, end=num_rows, alpha=alpha),
                                     {'type': 'data_bar', 'bar_color': '#63C384', 'bar_direction': 'right'})
    # Open, Base, Close, Prev Close
    for alpha in ['O', 'R', 'S', 'T', 'U', 'V']:
        worksheet.conditional_format('{alpha}{start}:{alpha}{end}'.format(start=2, end=num_rows, alpha=alpha),
                                     {'type': '3_color_scale'})
    for i in range(0, len(df)):
        diff52w = df['52wDiff'].iloc[i].round(2)
        worksheet.write('E{}'.format(i + 2), diff52w,
                        workbook.add_format({'num_format': '#,###',
                                             'fg_color': cost_color_codes[cost_per_range(diff52w)],
                                             'border': 1}))
        # Returns
        current_returns = df['CurRet%'].iloc[i]
        worksheet.write('L{}'.format(i + 2), current_returns,
                        workbook.add_format({'num_format': '0.00%', 'bold': True,
                                             'border': 1}))
        mid_returns = df['MidRet%'].iloc[i]
        worksheet.write('M{}'.format(i + 2), mid_returns,
                        workbook.add_format({'num_format': '0.00%', 'bold': True,
                                             'border': 1}))
        low_returns = df['LowRet%'].iloc[i]
        worksheet.write('N{}'.format(i + 2), low_returns,
                        workbook.add_format({'num_format': '0.00%', 'bold': True,
                                             'border': 1}))
    if close:
        workbook.close()


if __name__ == '__main__':
    file_masterdata = os.path.join(path_output, file_masterdata)
    file_insight = os.path.join(path_output, file_insight)
    df = pd.read_excel(file_masterdata)
    output_file = file_insight
    if os.path.exists(output_file):
        os.remove(output_file)

    # Main Call
    sheet_name = 'Nifty-All'
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter',
                            datetime_format='mmm dd yyyy',
                            date_format='mmmm dd yyyy',
                            options={'strings_to_urls': False, 'strings_to_formulas': False})

    df = modify_columns(df)
    df.drop_duplicates(subset=['Series', 'Symbol'], inplace=True)
    df = df[df['Series'] == 'EQ']
    # Nifty-Index
    df['Nifty-Index'] = df['Nifty-Index'].apply(parse_nifty)
    # Differences
    df['dayDiff'] = df['dayHigh'] - df['dayLow']
    df['52wDiff'] = df['52wHigh'] - df['52wLow']
    # Date Difference
    df['52wDateDiff'] = df['52wHighDate'] - df['52wLowDate']
    # 52 week Middle Median Point = (high + Low) / 2
    df['52wMid'] = (df['52wHigh'] + df['52wLow']) / 2
    # 52 week Mid Analysis
    df['avg-52Mid'] = df['avgPrice'] - df['52wMid']
    df['MidRet%'] = df['avg-52Mid'] / df['52wMid']
    # 52 week Low Analysis
    df['avg-52Low'] = df['avgPrice'] - df['52wLow']
    df['CurRet%'] = df['avg-52Low'] / df['52wLow']
    # 52w Low Return
    df['LowRet%'] = df['52wDiff'] / df['52wLow']
    # change percent today & previous day
    df['change'] = df['change'].apply(convert_float)
    df['pChange'] = df['pChange'].apply(convert_float)
    df = reorder_columns(df, extra_col=False)
    # Sort Parameter
    df.sort_values(by=['52wLowDate'], ascending=False, inplace=True)
    # Writing & Formatting Excel
    df.to_excel(writer, sheet_name=sheet_name, index=False, encoding='utf-8')
    format_nifty(writer, sheet_name=sheet_name, df=df)
    sectors = dict(df['Industry'].value_counts())
    print('Total Nifty listed company - ', df['Industry'].count())
    len_sector = df['Industry'].count()
    for sector in sectors.keys():
        percent = round((sectors[sector]/len_sector) * 100, 2)
        print('Writing Sector - ', sector, sectors[sector], 'Percent - ', percent)
        tmp_df = df[df['Industry'] == sector].copy()
        tmp_df.sort_values(by=['52wLowDate'], ascending=False, inplace=True)
        sheet_name = sector + '-' + str(sectors[sector])
        tmp_df.to_excel(writer, sheet_name=sheet_name, index=False)
        format_nifty(writer, sheet_name=sheet_name, df=tmp_df)
    else:
        format_nifty(writer, sheet_name=sheet_name, df=tmp_df, close=True)
    print(sectors)
    print('Done writing file.')

    print('Investment for buying from all sectors :-')
    amount = 200000
    print('Total Capital is Rs.', amount)
    for sector in sectors.keys():
        percent = round((sectors[sector]/len_sector) * 100, 2)
        buy_amount = round((amount * percent)/100, 2)
        print('{}({}) - {}% - Rs.{}'.format(sector, sectors[sector], percent, buy_amount, amount))

# 1 year Returns (high/low * 100) 4.5x
# current returns (current/low * 100) 0.25 x
# 52 W low & high indicator
