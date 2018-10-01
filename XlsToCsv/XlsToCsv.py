# -*- coding: utf-8 -*-

import xlrd
import io

rb = xlrd.open_workbook('1.xlsm')
csv = []
for i in range(rb.nsheets):
    sheet = rb.sheet_by_index(i)
    for j in range(sheet.nrows):
        row = sheet.row_values(j)
        row_csv = []
        for cell in row:
            row_csv.append(str(cell))
        csv.append(';'.join(row_csv) + '\n')

with io.open('out.csv', 'w', encoding='utf8') as f:
    f.writelines(csv)

