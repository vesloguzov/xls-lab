# -*- coding: UTF-8 -*-
import sys
import random
import datetime
import time

from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.reader import reader
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.styles import *
from lab_1_create_template import create_template


reload(sys)
sys.setdefaultencoding('utf8')

# Создание шаблона
wb = Workbook()
ws = wb.active
employees = ["Иванов И.М.", "Коробова П.Н", "Морозов И.Р.", "Ромашова П.Т.", "Петров Г.Т.", "Смирнов С.И.", "Соколова О.С."]
positions = ["Директор", "Менеджер", "Бухгалтер", "Зам. директора", "Секетарь", "Водитель", "Строитель"]
ws = create_template(ws, employees, positions)
wb.save('lab1_template.xlsx')



print ws.cell(row=7+7, column=2).coordinate

# print "D7: ", ws[interesCell].value
# print "D7: ", ws[interesCell].coordinate
# print "D7: ", ws[interesCell].column

print 'saved'