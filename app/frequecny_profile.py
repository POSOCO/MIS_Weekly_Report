from openpyxl import Workbook
from openpyxl import load_workbook
import pandas as pd 

wb =load_workbook(filename = '../Files/Frequency.xlsx')
ws = wb.active


frequency_cell = ws['B2:B8641']
frequency_value = []

for row in frequency_cell:
    for col in row:
        frequency_value.append(col.value)


print(max(frequency_value))

print(min(frequency_value))


def avg_frequency(lst):
    return (sum(lst)/len(lst))


print(round(avg_frequency(frequency_value),2))