import os
#import glob
#import re
#import openpyxl
import xlrd
from xlrd import open_workbook
import csv
f = '1033_files.xlsx'
workbook = open_workbook(f)
all_file = open('all_1033.csv', 'wb')
writer = csv.writer(all_file, quoting = csv.QUOTE_ALL)
header = ["State","Agency Name","NSN","Item Name","Quantity","UI","Ship Date"]
writer.writerow(header)
for s in workbook.sheets():
	for rownum in xrange(s.nrows):
		new_row = s.row_values(rownum)
		if new_row[0] != header[0]:
			writer.writerow(new_row)
all_file.close()

	

#######################
##  Failed attempt ####
#######################

# for sheet in workbook.worksheets:
#     filename = sheet.title.replace(" ", "_") + '.xlsx'
#     wb = openpyxl.Workbook()
#     to_convert = workbook.get_sheet_by_name(sheet.title)
#     print to_convert
#     wb.active = to_convert
#     wb.save(filename = filename)

# sheet_list = glob.glob("*.xlsx")

# for sheet in sheet_list:
# 	new = sheet.strip('\.xlsx')
# 	print new
# 	os.system('in2csv %s > %s.csv ' % (sheet, new))

# os.system('csvstack *.csv')
