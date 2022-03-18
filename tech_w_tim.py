from openpyxl import Workbook, load_workbook

wb = load_workbook('Book1.xlsx')

# get the sheets
ws = wb.active
print(ws['A1'])
