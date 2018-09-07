# *-* coding: utf-8 *-*

"""
    Combine all nifty documents
"""

import os
import sys
import pandas as pd
import json
import time
import random

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
sys.path.insert(1, ROOT_DIR)
for path, dirs, files in os.walk(ROOT_DIR):
    if '__pycache__' in path:
        continue
    sys.path.insert(1, path)

from etc.config import path_files, path_output, path_tmp, index_score, file_masterdata
from lib.util import OsUtil, DateUtil
from etc.config import get_quote
from lib.scrap import _request

global idx
idx = 0


def get_quote_data(x):
    global idx
    symbol = x['Symbol']
    idx += 1
    tmp_file = OsUtil.join_path(path_tmp, DateUtil.get_date('%Y_%m_%d') + '_' + symbol + '.json')
    if not os.path.isfile(tmp_file):
        time.sleep(random.randrange(1, 7, 1) / 10.0)
        r = _request(get_quote.format(symbol.strip()))
        _data = r.html.xpath('//*[@id="responseDiv"]', first=True).text
        if not _data:
            return dict()
        source_data = json.loads(_data)
        print('[{}]'.format(idx), '**** ', symbol, r.url)
        if not source_data['data']:
            print('%% ERROR %%  ', symbol)
            return dict()
        else:
            with open(tmp_file, 'w') as fp:
                json.dump(source_data, fp, indent=4)
    else:
        print('[{}]'.format(idx), '**** ', symbol)
        with open(tmp_file) as fp:
            source_data = json.load(fp)
    print('{} - {}'.format(symbol, source_data))
    source_data = source_data['data'][0]
    keys = list(source_data.keys())
    keys.sort()
    for key in keys:
        _value = source_data[key]
        if ',' in str(_value):
            _value = _value.replace(',', '')
        x[key] = _value
    return x


def _concat(x):
    _tag = list(x['TAG'])
    _tag.sort()
    return ', '.join(_tag)


def _tag_score(x):
    return len(x.split(','))


def _algo_score(x):
    x = x.split(',')
    _score = 0
    for _tag in x:
        _tag = _tag.strip()
        _score += index_score[_tag]
    return _score


if __name__ == '__main__':
    for count in range(5):
        try:
            group_list = ['Company Name', 'Industry', 'Symbol', 'Series', 'ISIN Code']
            file_name = OsUtil.list_files_in_path(path_files)
            file_output = os.path.join(path_output, file_masterdata)
            tags = dict()
            df = []
            for _file in file_name:
                tag = _file.replace('ind_', '').replace('list.csv', '')
                tags[tag] = 0
                _df = pd.read_csv(os.path.join(path_files, _file))
                _df['TAG'] = tag
                df.append(_df)

            df = pd.concat(df, ignore_index=True)
            df = df.groupby(group_list).apply(_concat).reset_index(name='TAGS')
            df['TAG_SCORE'] = df['TAGS'].apply(_tag_score)
            df['ALGO_SCORE'] = df['TAGS'].apply(_algo_score)
            df = df.apply(get_quote_data, axis=1)
            # df.rename(index=str, columns=rename_columns, inplace=True)  # rename
            columns = list(df.columns)
            columns.sort()
            for column in columns:
                df[column] = pd.to_numeric(df[column], errors='ignore')
                df[column] = pd.to_datetime(df[column], format='%d-%b-%y', errors='ignore', dayfirst=True)
            # df = df[align_columns]
            df.sort_values(by=['dayLow'], inplace=True, ascending=True)  # sort
            writer = pd.ExcelWriter(file_output, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='All_Nifty')
            workbook = writer.book
            worksheet = writer.sheets['All_Nifty']
            # Add a header format.
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'align': 'center',
                'fg_color': '#D7E4BC',
                'border': 1})

            # Write the column headers with the defined format.
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            format1 = workbook.add_format({'num_format': '0.00'})
            worksheet.set_zoom(99)
            worksheet.set_column('A1:BU1', 17, format1)  # Adds formatting to column C
            writer.save()
            print(list(df.columns))
            break
        except:
            print('Error running again - {}'.format(count+1))
            continue

# dayDiff, high-Low, max change, year, month, day,
