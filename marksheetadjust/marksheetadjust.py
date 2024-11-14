"""
Populate demo data into marksheet files in a directory using StudID for PyCon AU
Not an actual use case in assignment marking
"""

import xlwings as xw
import os

sheet_cond = "Ass1.xlsx"
adjust = {
	0 : 0,
	1 : 2,
	2 : 3,
	3 : 4,
	4 : 5,
	5 : 5}
directory = os.fsencode(".")

    
for file in os.listdir(directory):
	filename = os.fsdecode(file)
	if filename.endswith(sheet_cond):
		print (filename)

		# Read xlsx
		wb = xw.Book(filename)
		app = xw.apps.active
		sht = xw.sheets.active
		sht.range('D7').value = adjust[sht.range('D7').value]
		sht.range('D8').value = adjust[sht.range('D8').value]
		wb.save()
		app.quit()

		
		
		
      
