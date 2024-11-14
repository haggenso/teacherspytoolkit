"""
Populate demo data into marksheet files in a directory using StudID for PyCon AU
Not an actual use case in assignment marking
"""

import xlwings as xw
import os

sheet_cond = "Ass1.xlsx"
directory = os.fsencode(".")

    
for file in os.listdir(directory):
	filename = os.fsdecode(file)
	if filename.endswith(sheet_cond):
		print (filename)

		# Read xlsx
		wb = xw.Book(filename)		
		app = xw.apps.active
		sht = xw.sheets.active

		StudID = str(sht.range('B2').value)
		sht.range('D5').value = int(StudID[0:2])
		sht.range('D6').value = int(StudID[2:4])
		sht.range('D7').value = int(StudID[4:5])
		sht.range('D8').value = int(StudID[5:6])
		wb.save()		
		app.quit()



		
		
      
