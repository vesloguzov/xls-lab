from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.layout import Layout, ManualLayout
import xlsxwriter
import xlwings as xw


wb2 = load_workbook('data\pattern.xlsx')
ws = wb2['Sheet']
for row in ws.rows:
    for cell in row:
        print(cell.value)