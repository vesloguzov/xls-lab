# -*- coding: UTF-8 -*-
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.reader import reader
from openpyxl.chart.layout import Layout, ManualLayout
# import xlsxwriter
# import xlwings as xw

reload(sys)
sys.setdefaultencoding('utf8')

wb2 = load_workbook('labWork10.xlsx')
ws = wb2['Фильтр']

# print ws._charts
# print wb2.get_sheet_names()
# print type(ws.auto_filter.ref)
# print ws.auto_filter
first_cell = ws.auto_filter.ref.split(':')[0]
cell_date_filter = ''
cell_cost_filter = ''
# print ws[first_cell].row
for Colfilter in ws.auto_filter.filterColumn:
    if Colfilter.filters is not None:
        cell_date_filter = ws[first_cell].column + str(Colfilter.colId)
        print type(cell_date_filter)
        for Colfilter2 in Colfilter.filters.dateGroupItem:
            print Colfilter2.year
    print Colfilter.colId #относительно ws.auto_filter.ref
    if Colfilter.customFilters is not None:
        # print Colfilter.colId
        for Colfilter1 in Colfilter.customFilters.customFilter:
            print Colfilter1.operator+": "+Colfilter1.val
            pass
