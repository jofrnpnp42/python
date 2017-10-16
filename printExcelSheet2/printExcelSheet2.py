#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
Excelファイル sample.xls の [申請書]シートの有効セルを表示する
+ シートの存在有無をチェックし、なければシートへのアクセスを試みないようにする
"""
import sys
import xlrd

# シートの有無をチェックする
def isExistSheet(book, sheet):
    for i in range(0, book.nsheets):
        name = book.sheet_by_index(i).name
        if name == sheet:
            return  True
    return  False

# シートの有効セルを全て表示する
def printSheet(sheet):
    for col in range(sheet.ncols):
        for row in range(sheet.nrows):
            print(sheet.cell(row, col).value)

def main():
    sheet   = u'申請書'
    book    = xlrd.open_workbook('sample.xls')
    isExist = isExistSheet(book, sheet)
    if isExist == True:
        printSheet(book.sheet_by_name(sheet))


if __name__ == '__main__':
    main()

