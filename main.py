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

wb2 = load_workbook('data/rabota 8.xlsx')
ws = wb2[wb2.get_sheet_names()[0]]

# print ws._charts
# print wb2.get_sheet_names()
# print type(ws.auto_filter.ref)
# # print ws.auto_filter
# first_cell = ws.auto_filter.ref.split(':')[0]
# cell_date_filter = ''
# cell_cost_filter = ''
# print ws[first_cell].row
# for Colfilter in ws.auto_filter.filterColumn:
#     if Colfilter.filters is not None:
#         cell_date_filter = ws[first_cell].column + str(Colfilter.colId)
#         print cell_date_filter
#         for Colfilter2 in Colfilter.filters.dateGroupItem:
#             print Colfilter2.year
#     print Colfilter.colId #относительно ws.auto_filter.ref
#     if Colfilter.customFilters is not None:
#         # print Colfilter.colId
#         for Colfilter1 in Colfilter.customFilters.customFilter:
#             print Colfilter1.operator+": "+Colfilter1.val
#             pass



def string_is_money_format(str):
    if '₽' in str:
        return True
    if '£' in str:
        return True
    if '$' in str:
        return True
    if '€' in str:
        return True
    return False


def range_is_money_format(range):
    rows = ws[range]
    for row in rows:
        for cell in row:
            if not string_is_money_format(cell.number_format):
                return False
    return True

print range_is_money_format('F3:F6')

# print "D7: ", ws['F3'].number_format
# print "D7: ", ws['F4'].number_format
# print "D7: ", ws['F5'].number_format
#
# print "D7: ", ws['F6'].number_format

# первый столбоц ширина 30
