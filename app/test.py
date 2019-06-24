from docx import Document
from docx.shared import Inches
from datetime import datetime
f = open('../Logs/Logger.txt',"a+")
f.write('The code is run at '+datetime.now().strftime('%Y-%m-%d %H:%M:%S')+ '\n')
document = Document()

table = document.add_table(rows = 10,cols = 11)
table.style = 'Table Grid'
table.autofit = True
header_cells = table.rows[0].cells  

header_cells[0].text = 'दिनांक'
header_cells[0].width = Inches(4.5)
header_cells[1].text = 'अिधकतम लक '
header_cells[2].text = 'नतम तालक'
header_cells[3].text = 'औसत'
header_cells[4].text = '% of time 49.9 से कम '
header_cells[5].text = '% of time 49.9 से 50.05 तक '
header_cells[6].text = '% of time 50.05 से उपर '
header_cells[7].text = 'Percentage of Time Frequency Remained outside IEGC Band'
header_cells[8].text = 'No’s of hours frequency remained outside IEGC Band '
header_cells[9].text = 'FDI=No. of hours out of IEGC band/24Hrs '
header_cells[10].text = 'Weekly FDI=No. of hours out of IEGC band/24*7Hrs '

for cell in table.columns[0].cells:
    cell.width = Inches(2.5)

document.save('demo.docx')