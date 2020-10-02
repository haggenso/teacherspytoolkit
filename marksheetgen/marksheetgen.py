from openpyxl import load_workbook
from openpyxl import Workbook
import pandas as pd

# Read Student Roster
studroster = pd.read_csv("classroster.csv")
print(studroster)

# Read xlsx
wb = load_workbook(filename = 'A118_eng.xlsx')
ws = wb.active

