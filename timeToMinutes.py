import openpyxl
import re
import os

os.chdir('/Users/darrick/Downloads/')
wb = openpyxl.load_workbook('Time-1.xlsx')
sheet = wb['Sheet1']

def timeToMinutes():
    rows = sheet.max_row + 1
    for i in range(2, rows):
        data = "A" + str(i)
        new = "C" + str(i)
        sheet[new].value = sheet[data].value[6:14]
    wb.save('Time2.xlsx')