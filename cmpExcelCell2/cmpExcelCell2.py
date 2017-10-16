#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
import datetime
import sys
import os
import fnmatch
import xlrd

# シートの有無をチェックする
def isExistSheet(book, sheet):
    for i in range(0, book.nsheets):
        name = book.sheet_by_index(i).name
        if name == sheet:
            return  True
    return  False

#
def walk_dir(d):
    for root, dirs, files in os.walk(d):
        yield root
        for file in files:
            yield os.path.join(root, file)


# 個数が一致しているかを検証する。
# 不一致の場合は output で指定したファイルにファイル名と個数を書き出す。
def assertEqItemNum( f, sheet, output ):
    with open(output, "a+") as fp:
        MaxRow = sheet.nrows
        #print("[{0}] MaxRow={1}".format(sheet.name, MaxRow))
        for row in range(2, MaxRow):
            name = sheet.cell(row, 1).value
            if name != "":
                inum = sheet.cell(row, 3).value
                onum = sheet.cell(row, 4).value
                if inum == onum:
                    print("#%s,%s,%d,%d" % (f, name, inum, onum))
                else:
                    fp.write("%s,%s,%d,%d\n" % (f, name, inum, onum))
                    print("%s,%s,%d,%d" % (f, name, inum, onum))

#
def main():
    sheet = u'概況'
    argv  = sys.argv
    argc  = len(argv)
    if(argc == 1):
        paths = '.'
    else:
        paths = argv

    now   = datetime.datetime.now()
    fmt_name = "mismatch_{0:%Y%m%d-%H%M%S}.csv".format(now)

    for p in paths:
        for f in walk_dir(p):
            if fnmatch.fnmatch(f, '*在庫管理*.xls*'):
                book = xlrd.open_workbook(f)
                isExist = isExistSheet(book, sheet)
                if isExist == True:
                    assertEqItemNum(f, book.sheet_by_name(sheet), fmt_name)


if __name__ == '__main__':
    main()
