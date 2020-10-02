"""
This program read in class roster from the file "classroster.csv" with the format:

StudID	StudName	Class
A001	Name1	A

The Excel template is in the file "marksheet_template.xlsx". 

"""

from openpyxl import load_workbook
from openpyxl import Workbook
import pandas as pd

# Read Student Roster
df = pd.read_csv("classroster.csv")
# print(df)

# Loop
for index, row in df.iterrows():
	# print (row['StudID'],row['StudName'],row['Class'])

	# Read xlsx
	wb = load_workbook(filename = 'marksheet_template.xlsx')
	ws = wb.active
	# print(ws)

	ws['B1']=row['StudName']
	ws['B2']=row['StudID']
	ws['B3']=row['Class']
	wb.save(row['StudID'] + "_" + row['StudName'] + "_" + "Ass1.xlsx")	
