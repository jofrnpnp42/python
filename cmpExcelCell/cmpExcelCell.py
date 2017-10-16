#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
Excelファイル sample.xls の [申請書]シートの郵便番号、住所を比較する。
+ シートの存在有無をチェックし、なければシートへのアクセスを試みないようにする
+「D13 vs D17」「F13 vs F17」「D14 vs D18」の比較をする
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

# 郵便番号と住所が一致していたら真を返す
def isEqualAddress( sheet ):
    oldP     = sheet.cell(12, 3).value  # 13行目、D列  (郵便番号 XXX-    )
    curP     = sheet.cell(16, 3).value  # 17行目、D列  (郵便番号 XXX-    )
    oldC     = sheet.cell(12, 5).value  # 13行目、F列  (郵便番号    -YYYY)
    curC     = sheet.cell(16, 5).value  # 17行目、F列  (郵便番号    -YYYY)
    old_addr = sheet.cell(13, 3).value  # 14行目、D列  (住所)
    cur_addr = sheet.cell(17, 3).value  # 18行目、D列  (住所)
    print("{0}-{2} vs {1}-{3}".format(oldP, curP, oldC, curC))
    print("旧住所: {0}".format(old_addr))
    print("現住所: {0}".format(cur_addr))
    if oldP == curP and oldC == curC:
        return True
    else:
        return False

#
def main():
    sheet   = u'申請書'
    book    = xlrd.open_workbook('sample.xls')
    isExist = isExistSheet(book, sheet)
    if isExist == True:
        isMatch = isEqualAddress(book.sheet_by_name(sheet))
        if isMatch == True:
            print("住所に変更はありません")
        else:
            print("住所が変わりました")


if __name__ == '__main__':
    main()

