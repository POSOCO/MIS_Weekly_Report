from openpyxl import Workbook
from openpyxl import load_workbook
import pandas as pd 

def avg_frequency(lst):
    return (sum(lst)/len(lst))


wb =load_workbook(filename = '../Files/Frequency.xlsx')
ws = wb.active


frequency_cell = ws['B2:B8641']
frequency_value = []

for row in frequency_cell:
    for col in row:
        frequency_value.append(col.value)


max_frequency = round(max(frequency_value),2)

min_frequency = round(min(frequency_value),2)

avg_frequency = round(avg_frequency(frequency_value),2) 


def band_duration_info():
    below_band = 0
    in_the_band = 0
    above_band = 0
    percent_value = 86.40
    for each in frequency_value:
        if each < 49.90:
            below_band += 1
        elif each >= 49.90 and each <= 50.05:
            in_the_band += 1
        else:
            above_band += 1

    band_duration_info = {
        'below_band':round(below_band/percent_value,2),
        'in_the_band':round(in_the_band/percent_value,2),
        'above_band':round(above_band/percent_value,2)
    }
    return band_duration_info

print(band_duration_info())