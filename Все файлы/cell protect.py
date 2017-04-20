from openpyxl import Workbook

wb = Workbook()

ws = wb.active

ws.protection.sheet = True

ws.protection.set_password('test')

cell = ws['A1']
cell.protection = Protection(locked=False) # default is True because a

wb.save(filename = 'simple.xlsx')