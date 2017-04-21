# -*- coding: UTF-8 -*-
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.reader import reader
from openpyxl.chart.layout import Layout, ManualLayout

reload(sys)
sys.setdefaultencoding('utf8')

def range_is_date_format(ws, range):
    rows = ws[range]
    for row in rows:
        for cell in row:
            if not cell.is_date:
                return {'status': False, 'message': 'Dates invalid'}
    return {'status': True, 'message': 'Dates valid'}

def range_is_money_rub_format(ws, range):
    rows = ws[range]
    for row in rows:
        for cell in row:
            if not '₽' in cell.number_format:
                return {'status': False, 'message': 'Money rub format invalid'}
    return {'status': True, 'message': 'Money rub format valid'}


def check_ranges_equal(ws_correct, correct_range, ws_student, student_range):
    correct_rows = ws_correct[correct_range]
    correct_list = []
    for row in correct_rows:
        for cell in row:
            correct_list.append(cell.value)

    student_rows = ws_student[student_range]
    student_list = []
    for row in student_rows:
        for cell in row:
            student_list.append(cell.value)

    print "correct_list", correct_list
    print "student_list", student_list

def formulas_is_equal(f1, f2):
    if f1 and f2:
        f1 = f1.replace(" ", "").lower().replace(".", ",")
        f2 = f2.replace(" ", "").lower().replace(".", ",")

        return f1 == f2

    else: return False

def check_cost(ws_student, student_range):
    student_rows = ws_student[student_range]
    for index, row in enumerate(student_rows):
        for cell in row:
            var1 = '=E'+str(index+3)+'*D'+str(index+3)
            var2 = '=D' + str(index + 3) + '*E' + str(index + 3)
            if formulas_is_equal(cell.value, var1) == False and  formulas_is_equal(cell.value, var2) == False:
                return False
    return True

def check_formats(student_ws):
    # Проверяем правильность форматирования дат поступления
    print "Формат дат поступления: ", range_is_date_format(student_ws, 'C3:C17')['status']

    # Проверяем правильность форматирования цены
    print "Формат цены товара: ", range_is_money_rub_format(student_ws, 'E3:E17')['status']

    # Проверяем правильность форматирования стоимости
    print "Формат суммарной стоимости: ", range_is_money_rub_format(student_ws, 'F3:F17')['status']

def check_answer(correct_wb, correct_wb_data_only, student_wb, student_wb_data_only, data):
    # try:
        # print student_wb.get_sheet_names()

    student_ws_1 = student_wb[student_wb.get_sheet_names()[0]]
    student_ws_2 = student_wb[student_wb.get_sheet_names()[1]]
    student_ws_3 = student_wb[student_wb.get_sheet_names()[2]]

    correct_ws_1 = correct_wb[correct_wb.get_sheet_names()[0]]
    correct_ws_2 = correct_wb[correct_wb.get_sheet_names()[1]]
    correct_ws_3 = correct_wb[correct_wb.get_sheet_names()[2]]

    ws_read_only_1 = student_wb_data_only[student_wb_data_only.get_sheet_names()[0]]
    ws_read_only_2 = student_wb_data_only[student_wb_data_only.get_sheet_names()[1]]
    ws_read_only_3 = student_wb_data_only[student_wb_data_only.get_sheet_names()[2]]
    cost_range = 'F3:F17'
    #check_ranges_equal(correct_ws_1, cost_range, student_ws_1, cost_range)
    if check_cost(student_ws_1, cost_range):
        print 'Стоимость заполнена'
        check_formats(student_ws_1)



    else: print 'Стоимость не заполнена'

    # print data
    # except:
    #     print 'Документ должен содержать три рабочих листа'