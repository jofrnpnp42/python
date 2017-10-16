#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
import datetime
import sys
import os
import fnmatch
import xlrd
import re

#
def walk_dir(d):
    for root, dirs, files in os.walk(d):
        yield root
        for file in files:
            yield os.path.join(root, file)


# 商品の在庫状況を表示する
def inventorySupply( f, sheet, output ):
    # with により close の呼び出し不要になる
    with open(output, "a+") as fp:
        for col in range(sheet.ncols):
            for row in range(sheet.nrows):
                #「商品名」を見つけたら、「No」がなくなるまでその列(col)を下に辿っていく
                itemName = sheet.cell(row, col).value
                if "商品名" == itemName:
                    for row in range(row, sheet.nrows):
                        idx  = sheet.cell(row, col-1).value
                        #「No」なし＝そのシートでの商品チェック終了
                        if idx == "":
                            return

                        itemName = sheet.cell(row, col).value   # 商品名
                        purNum   = sheet.cell(row, col+2).value # 購入数(purchase)
                        soldNum  = sheet.cell(row, col+3).value # 販売数(sold)
                        # print("{0},{1},{2},{3}".format(f, itemName, purNum, soldNum))

                        # 空欄(記入漏れ)があれば次項目に進む
                        if itemName == "" or purNum == "" or soldNum == "":
                            fp.write("[Empty] %s,%d,%d\n".format(itemName, purNum, soldNum))
                            continue

                        if purNum == soldNum:
                            fp.write("[SoldOut] %s,%d,%d\n".format(itemName, purNum, soldNum))
                        elif purNum > soldNum:
                            fp.write("[Rest] %s,%d,%d\n".format(itemName, purNum, soldNum))
                        else:
                            fp.write("[Illegal] %s,%d,%d\n".format(itemName, purNum, soldNum))


#
def main():
    argv  = sys.argv
    argc  = len(argv)
    if(argc == 1):
        paths = '.'
    else:
        paths = argv

    now   = datetime.datetime.now()
    ofile = "result_{0:%Y%m%d-%H%M%S}.csv".format(now)

    for p in paths:
        for f in walk_dir(p):
            if fnmatch.fnmatch(f, '*在庫管理*.xls*'):
                book = xlrd.open_workbook(f)
                for i in range(0, book.nsheets):
                    sheetName = book.sheet_by_index(i).name
                    print(sheetName)
                    if re.match('201[0-9][0-9]{2}', sheetName):  # シート名が "201610' など
                        sheet = book.sheet_by_index(i)
                        inventorySupply(f, sheet, ofile)


if __name__ == '__main__':
    main()
