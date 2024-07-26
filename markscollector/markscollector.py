"""
Collect student details and marks from all the *.xlsx files in a directory
Using xlwings library that can read calculated results from formulae
"""

import xlwings as xw
import pandas as pd
import os

df = pd.DataFrame(columns=['StudID','StudName','Class','Mark'])
i = 0

directory = os.fsencode(".")
    
for file in os.listdir(directory):
	filename = os.fsdecode(file)
	if filename.endswith(".xlsx"):
		print (filename)

		# Read xlsx
		wb = xw.Book(filename)
        #app = xw.apps.active
		sht = xw.sheets.active
		#print (ws['B1'].value)
		
		row = []
		row.append(sht.range('B1').value)
		row.append(sht.range('B2').value)
		row.append(sht.range('B3').value)
		row.append(sht.range('B10').value)
        #app.quit()
		df.loc[i] = row
		i += 1
	
#print(df)
df.to_csv('StudMark.csv')
		
		
      
