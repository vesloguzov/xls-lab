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

print type(ws.merged_cell_ranges[0])

# my_range = wb2.defined_names[ws.merged_cell_ranges[0]]
# dests = my_range.destinations # returns a generator of (worksheet title, cell range) tuples
#
# cells = []
# for title, coord in dests:

print ws.auto_filter.ref

interesCell = 'G4'

print "D7: ", ws[interesCell].value
print "D7: ", ws[interesCell].coordinate
print "D7: ", ws[interesCell].column
print "D7: ", ws[interesCell].base_date
print "D7: ", ws[interesCell].guess_types
print "D7: ", ws[interesCell].internal_value
print "D7: ", ws[interesCell].is_date
print "D7: ", ws[interesCell].number_format

# print "B7: ", ws['B7'].value
# print "B7: ", ws['B7'].coordinate
# print "B7: ", ws['B7'].column
# print "B7: ", ws['B7'].base_date
# print "B7: ", ws['B7'].guess_types
# print "B7: ", ws['B7'].internal_value
# print "B7: ", ws['B7'].is_date
# print "B7: ", ws['B7'].number_format
