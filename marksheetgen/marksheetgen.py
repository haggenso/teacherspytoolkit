"""
This program read in class roster from the file "classroster.csv" with the format:

StudID	StudName	Class
A001	Name1	A

The Excel template is in the file "marksheet_template.xlsx".
You can also switch between sheetw with the statement
ws = wb[sheets[0]]

Ref: https://stackoverflow.com/questions/45756731/how-to-switch-between-sheets-in-excel-openpyxl-python-to-make-changes?rq=1

"""

from openpyxl import load_workbook
import pandas as pd

# Read Student Roster
df = pd.read_csv("classroster.csv")
# print(df)

# Loop
for index, row in df.iterrows():
	# print (row['StudID'],row['StudName'],row['Class'])

	# Read xlsx
	wb = load_workbook(filename = 'marksheet_template.xlsx')
	sheets = wb.sheetnames
	ws = wb[sheets[0]]
	ws = wb.active
	# print(ws)

	ws['B1']=row['StudName']
	ws['B2']=row['StudID']
	ws['B3']=row['Class']
	wb.save(row['StudID'] + "_" + row['StudName'] + "_" + "Ass1.xlsx")	
