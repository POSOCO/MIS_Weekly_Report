from openpyxl import Workbook
from openpyxl import load_workbook
import glob
from datetime import datetime

f = open('../Logs/Logger.txt',"a+")
f.write('The code is run at Vol Profile Script '+datetime.now().strftime('%Y-%m-%d %H:%M:%S')+ '\n')

files = glob.glob("../Files/vol/*.xlsx")

vol_station_list_400KV = {'Bhopal','Khandwa','Itarsi','Damoh','Nagda','Indore','Gwalior','Raipur','Raigarh','Bhilai','Wardha','Dhule','Parli','Boisar','Kalwa','Karad','Asoj','Dehgam','Kasor',
'Jetpur','Amreli','Vapi'}
vol_station_list_765KV = {'Sipat','Seoni','Wardh','Bina','Indore','Sasan','Satna','Tamnar','Kotra','Vododara','Durg','Gwalior'}


for file in files:
    wb = load_workbook(filename = file)
    ws = wb.active
    for eachStation in vol_station_list_400KV:
        vol_cell = ws['C4']