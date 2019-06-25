#import vol_profile
from docx import Document
from datetime import datetime
from docx.shared import Pt

f = open('../Logs/Logger.txt',"a+")
f.write('The code is run at Vol Doc Script '+datetime.now().strftime('%Y-%m-%d %H:%M:%S')+ '\n')


doc = Document('../freq files/Format File Vol.docx')

table1 = doc.tables[0]
table2 = doc.tables[1]
table3 = doc.tables[2]
table4 = doc.tables[3]

style = doc.styles['Normal']
font = style.font
font.name = "Times New Roman"
font.size = Pt(9.5)
font.bold = True


table1Staions = {'Bhopal':'P','Khandwa':'G','Itarsi':'F','Damoh':'AZ','Nagda':'O','Indore':'N','Gwalior':'AS','Raipur':'BL','Raigarh':'BK'}
table2Stations = {'Bhilai':'AQ','Wardha':'AV','Dhule':'AE','Parli':'AC','Boisar':'CA','Kalwa':'V','Karad':'AA','Asoj':'AH','Dehgam':'L'}
table3Stations= {'Kasor':'AO','Jetpur':'AJ','Amreli':'BQ','Vapi':'BZ','Sipat':'AX','Seoni':'AW','Wardh':'BD','Bina':'BC','Indore':'BF'}
table4Stations = {'Sasan':'BH','Satna':'BG','Tamnar':'BW','Kotra':'BX','Vododara':'CF','Durg':'BY','Gwalior':'BB'}

# for row in range(7):
#     for col in range(9):
#         table1.cell(row,col)

table1.cell(2,2).text = '24'

doc.save('../freq files/Format File Vol1.docx')
    