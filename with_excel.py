#coding:utf-8

import xlrd

book = xlrd.open_workbook("DS4 connection.xlsx")
sheet = book.sheet_by_index(0)

for i in range(sheet.nrows):
    r = sheet.row_values(i)
    print(type(r[6]))
    print(r)


