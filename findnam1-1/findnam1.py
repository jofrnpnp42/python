#!/usr/bin/env python3

# https://docs.python.jp/3.3/library/csv.html
# https://docs.python.org/3.3/library/csv.html

import  csv

with open('meibo.csv', 'r', newline='') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')
    for row in spamreader:
        print(','.join(row))

