from openpyxl import Workbook
from openpyxl import load_workbook
import pandas as pd 

wb =load_workbook(filename = '../Files/Frequency.xlsx')
ws = wb.active

frequency_col_range = ws['B']
frequency_row_range = ws[2:81]
frequency_value = []

for row in frequency_row_range:
    for col in frequency_col_range:
        print(col.value)

# df = pd.DataFrame(frequency_value)

# print(max(frequency_value))