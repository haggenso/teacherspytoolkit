"""
This program read in each of the *.xlsx file in the current directory
and export them to pdf format in the subdirectory specified by the 
variable pdfrel

You can specify the sheet to print by changing the index of this statement
sheet = wb.sheets[0]

Ref:
https://stackoverflow.com/questions/57724345/print-excel-to-pdf-with-xlwings
https://analystcave.com/vba-excel-to-pdf-exporter-single-batch-export/
"""

import xlwings as xw
import os
import re
import time

pdfrel = "pdf"
directory = os.fsencode(".")
curdir = os.getcwd()
pdfabs = os.path.join(curdir, pdfrel)
if not os.path.exists(pdfabs):
	os.makedirs(pdfabs)
    
for file in os.listdir(directory):
	filename = os.fsdecode(file)
	if filename.endswith(".xlsx"):
		print (filename)

		# Read xlsx
		wb = xw.Book(filename)
		# Select Sheet
		sheet = wb.sheets[0]
		sheet.activate()
		
		filepdf = re.sub('xlsx$','pdf', filename)
		filepdf = os.path.join(pdfabs, filepdf)
		#print (filepdf)
        #You can use pywin32 api to print as well
		#wb.api.ExportAsFixedFormat(0, filepdf)
		sheet.to_pdf(path=filepdf)        
        
		time.sleep(0.5)
		wb.app.quit()
		time.sleep(0.5)
	
#print(df)