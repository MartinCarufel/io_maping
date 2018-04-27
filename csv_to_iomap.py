#coding:utf-8

import csv
import xlrd
import tkinter
from tkinter import filedialog


#file = open('DS4 connection.csv')
#csv_line = csv.reader(file)

root = tkinter.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()

book = xlrd.open_workbook(file_path)
sheet = book.sheet_by_index(0)
output_file = open('DS4-Analog Harness.xml','w')
x_line = []

output_file.write("<Map>\n")
output_file.write("\t<Devices>\n")

output_file.write("\t\t<Device model=\"8223\" deviceNo=\"1\">\n")


for row in range(sheet.nrows):
    #print(sheet.row_values(row, 6, 7))
    row_list = sheet.row_values(row)
    if row_list [6] == 'SeaDAC ISO 8223':
        output_file.write(
            "\t\t\t<Port name=\"{}\", type=\"{}\", portNo=\"{}\", comment=\"{}-{} {} {}\", negateValue=\"{}\"/>\n"
                .format(row_list[10],row_list[7], int(row_list[9]),row_list[0],row_list[4],row_list[3],row_list[5],
                row_list[8]))
output_file.write("\t\t</Devices>\n")


#file.close()
output_file.write("\t\t<Device model=\"8224\" deviceNo=\"2\">\n")
file = open(file_path)
csv_line = csv.reader(file)

for row in range(sheet.nrows):
    row_list = sheet.row_values(row)
    if row_list [6] == "SeaDAC ISO 8224":
        output_file.write(
            "\t\t\t<Port name=\"{}\", type=\"{}\", portNo=\"{}\", comment=\"{}-{} {} {}\", negateValue=\"{}\"/>\n".format(
                row_list[10], row_list[7], int(row_list[9]), row_list[0], row_list[4], row_list[3], row_list[5],
                row_list[8]))

output_file.write("\t\t</Devices>\n")

output_file.write("\t</Devices>\n")
output_file.write("\t<Groups />\n")
output_file.write("\t<BoolAnalogs/>\n")
output_file.write("</Map>")

#file.close()
output_file.close()