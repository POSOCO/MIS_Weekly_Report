from docx import Document
from docx.shared import Inches
from datetime import datetime
f = open('../Logs/Logger.txt',"a+")
f.write('The code is run at Test Script '+datetime.now().strftime('%Y-%m-%d %H:%M:%S')+ '\n')

alist = {'b':4,'a':3}

for each in alist:
    print(alist[each])