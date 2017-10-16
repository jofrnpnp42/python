#!/usr/bin/env python3
# -*- coding: utf-8 -*- 
"""
引数で指定したパス以下のファイル・ディレクトリをすべて表示する
"""
import  os
import  sys

def walk_dir(d):
    for root, dirs, files in os.walk(d):
        yield root
        for file in files:
            yield os.path.join(root, file)

def main():
    argv = sys.argv
    argc = len(argv)
    if(argc == 1):
        paths = '.'
    else:
        paths = argv
    for p in paths:
        for f in walk_dir(p):
            print(f)

if __name__ == '__main__':
    main()

