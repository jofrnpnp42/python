#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
Excelファイル sample.xls の [申請書]シートの有効セルを表示する
"""
import sys
import xlrd

book = xlrd.open_workbook('sample.xls')

# #### シート名を表示
# ### 最初のシート名
# print(book.sheet_by_index(0).name)
# ### "申請書"シート名
# print(book.sheet_by_name(u'申請書').name)

### "申請書"シートの有効行・列数を表示する
sheet = book.sheet_by_name(u'申請書') 
print(sheet.nrows)
print(sheet.ncols)

for col in range(sheet.ncols):
    for row in range(sheet.nrows):
        print(sheet.cell(row, col).value)

