#!/usr/bin/env python

import xlrd
import csv

def csv_from_excel():

    wb = xlrd.open_workbook('Sched.xlsx')
    sh = wb.sheet_by_index(0)
    your_csv_file = open('1.csv', 'wb')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in xrange(sh.nrows):
        wr.writerow([unicode(entry).encode("utf-8") for entry in sh.row_values(rownum)])

    your_csv_file.close()


def rm_hdr_lines_from_sched():
    #the sched spreadsheet has some bogus instruction lines that aren't 
    # necessary for our purposes - please nuke them for orbit
    with open('1.csv', 'r') as fin:
        data = fin.read().splitlines(True)
    with open('1.csv', 'w') as fout:
        fout.writelines(data[8:])


csv_from_excel()
rm_hdr_lines_from_sched()

