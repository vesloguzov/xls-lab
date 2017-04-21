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
from lab_1_check_answer import check_answer

reload(sys)
sys.setdefaultencoding('utf8')

# Создание шаблона
template_wb = Workbook()
template_ws = template_wb.active
employees = ["Иванов И.М.", "Коробова П.Н", "Морозов И.Р.", "Петров Г.Т.", "Ромашова П.Т.", "Смирнов С.И.", "Соколова О.С."]


positions = ["Директор", "Менеджер", "Бухгалтер", "Зам. директора", "Секетарь", "Водитель", "Строитель"]
dollar_rate = 48 #round(random.uniform(30, 60), 1)
template_ws = create_template(template_ws, employees, positions, dollar_rate)
template_wb.save('lab1_template.xlsx')


# Проверка
student_wb =  load_workbook('lab1_correct.xlsx')
student_wb_data_only =  load_workbook('lab1_correct.xlsx', data_only=True)

result = check_answer(student_wb, student_wb_data_only, employees, dollar_rate)


