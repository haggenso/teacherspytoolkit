# openpyxl can only read the value of very simple calcuated cell
# Even =SUM(...) does not work :'(
# Ref: https://stackoverflow.com/questions/23350581/openpyxl-1-8-5-reading-the-result-of-a-formula-typed-in-a-cell-using-openpyxl

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
		
		
      
