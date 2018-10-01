# -*- coding: utf-8 -*-

import xlrd
import io

def parse_cell(s, i):

    if not isinstance(s, float):
        return str.replace(str(s), ';', ',')
    precision = 0 if i == 0 else 2
    return ("{0:."+str(precision)+"f}").format(s)

rb = xlrd.open_workbook('1.xlsm')
csv = []
for i in range(rb.nsheets):
    sheet = rb.sheet_by_index(i)
    for j in range(sheet.nrows):
        row = sheet.row_values(j)
        row_csv = [parse_cell(row[k], k) for k in range(len(row))]   
        csv.append(';'.join(row_csv) + '\n')

with io.open('out.csv', 'w', encoding='utf8') as f:
    f.writelines(csv)