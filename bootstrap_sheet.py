import os

import openpyxl

os.chdir('/Users/darrick/Library/Mobile Documents/com~apple~CloudDocs/Coding_Dojo/Cheat_Sheets')
wb = openpyxl.load_workbook('bootstrap4_snippets.xlsx')
sheet = wb['Sheet1']

def moveData():
    column_offset = 2
    rowoffset = -68
    for i in range(70, 4040):
        if i % 70 == 0 and i != 70:
            column_offset += 2
            rowoffset -= 70
            
        sheet._move_cell(row=i, column=1, row_offset=rowoffset, col_offset = column_offset)
        sheet._move_cell(row=i, column=2, row_offset=rowoffset, col_offset = column_offset)

    wb.save('bootstrap4_snippets_1.xlsx')
