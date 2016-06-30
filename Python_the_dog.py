#!/usr/bin/env python

import xlrd
import csv

def csv_from_excel():

    wb = xlrd.open_workbook('Sched.xlsx')
    sh = wb.sheet_by_index(0)
    your_csv_file = open('your_csv_file.csv', 'wb')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in xrange(sh.nrows):
        wr.writerow([unicode(entry).encode("utf-8") for entry in sh.row_values(rownum)])

    your_csv_file.close()

csv_from_excel()
