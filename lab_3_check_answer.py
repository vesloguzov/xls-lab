# -*- coding: UTF-8 -*-
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.reader import reader
from openpyxl.chart.layout import Layout, ManualLayout
from utils import range_is_date_format, range_is_money_rub_format, formulas_is_equal

reload(sys)
sys.setdefaultencoding('utf8')

def check_ranges_equal(ws_correct, ws_student, range):
    correct_rows = ws_correct[range]
    correct_list = []
    for row in correct_rows:
        for cell in row:
            try:
                correct_list.append(round(float(cell.value), 2))
            except:
                correct_list.append(cell.value)

    student_rows = ws_student[range]
    student_list = []
    for row in student_rows:
        for cell in row:
            try:
                student_list.append(round(float(cell.value), 2))
            except:
                student_list.append(cell.value)

    return correct_list == student_list

def check_cost(ws_student, student_range):
    student_rows = ws_student[student_range]
    for index, row in enumerate(student_rows):
        for cell in row:
            var1 = '=E'+str(index+3)+'*D'+str(index+3)
            var2 = '=D' + str(index + 3) + '*E' + str(index + 3)
            if formulas_is_equal(cell.value, var1) == False and  formulas_is_equal(cell.value, var2) == False:
                return False
    return True

def check_results(correct_ws, student_ws):
    results_range = 'A2:E26'
    check_vals = check_ranges_equal(correct_ws, student_ws, results_range)
    check_rows = [4,6,8, 11, 15, 19, 21, 25, 26]
    for r in check_rows:
        if check_ranges_equal(correct_ws, student_ws, 'A'+str(r)+':F'+str(r)) == False:
            return False
    if check_vals:
        return True
    return False

def check_sorting(correct_ws, student_ws):
    range = 'A2:E17'
    return check_ranges_equal(correct_ws, student_ws, range)

def check_formats(student_ws):
    # Проверяем правильность форматирования дат поступления
    print "Формат дат поступления: ", range_is_date_format(student_ws, 'C3:C17')['status']

    # Проверяем правильность форматирования цены
    print "Формат цены товара: ", range_is_money_rub_format(student_ws, 'E3:E17')['status']

    # Проверяем правильность форматирования стоимости
    print "Формат суммарной стоимости: ", range_is_money_rub_format(student_ws, 'F3:F17')['status']

def get_date_custom_filters(ws):
    filters = {}
    filters['year'] = {}
    filters['year']['type'] = ''
    filters['year']['column'] = ''

    filters['custom'] = {}

    filters['custom']['column'] = ''
    filters['custom']['greaterThan'] = ''
    filters['custom']['lessThan'] = ''
    filters['range'] = ws.auto_filter.ref

    for Colfilter in ws.auto_filter.filterColumn:

        if Colfilter.dynamicFilter is not None:
            filters['year']['column'] = float(Colfilter.colId)
            filters['year']['type'] = Colfilter.dynamicFilter.type

        if Colfilter.customFilters is not None:
            for Colfilter1 in Colfilter.customFilters.customFilter:
                filters['custom']['column'] = float(Colfilter.colId)
                if Colfilter1.operator == 'greaterThan':
                    filters['custom']['greaterThan'] = float(Colfilter1.val)
                if Colfilter1.operator == 'lessThan':
                    filters['custom']['lessThan'] = float(Colfilter1.val)
    return filters

def check_filters(correct_ws, student_ws):

    is_data = check_ranges_equal(correct_ws, student_ws, 'D2:E17')
    is_filters = get_date_custom_filters(correct_ws) == get_date_custom_filters(student_ws)

    return is_data and is_filters

def check_answer(correct_wb, correct_wb_data_only, student_wb, student_wb_data_only, data):

    response_msg = {}

    if (len(student_wb.get_sheet_names()) == 3):

        student_ws_1 = student_wb[student_wb.get_sheet_names()[0]]
        student_ws_2 = student_wb[student_wb.get_sheet_names()[1]]
        student_ws_3 = student_wb[student_wb.get_sheet_names()[2]]

        correct_ws_1 = correct_wb[correct_wb.get_sheet_names()[0]]
        correct_ws_2 = correct_wb[correct_wb.get_sheet_names()[1]]
        correct_ws_3 = correct_wb[correct_wb.get_sheet_names()[2]]

        # student_ws_read_only_1 = student_wb_data_only[student_wb_data_only.get_sheet_names()[0]]
        # student_ws_read_only_2 = student_wb_data_only[student_wb_data_only.get_sheet_names()[1]]
        # student_ws_read_only_3 = student_wb_data_only[student_wb_data_only.get_sheet_names()[2]]
        #
        # correct_ws_read_only_1 = correct_wb_data_only[correct_wb_data_only.get_sheet_names()[0]]
        # correct_ws_read_only_2 = correct_wb_data_only[correct_wb_data_only.get_sheet_names()[1]]
        # correct_ws_read_only_3 = correct_wb_data_only[correct_wb_data_only.get_sheet_names()[2]]

        cost_range = 'F3:F17'

        if check_cost(student_ws_1, cost_range):
            print 'Стоимость заполнена'
            check_formats(student_ws_1)

            print "Сортировка выполнена верно: ", check_sorting(correct_ws_1, student_ws_1)

            print 'Лист итогов создан верно: ', check_results(correct_ws_2, student_ws_2)

            print 'Фильтры применены: ', check_filters(correct_ws_3, student_ws_3)


        else: print 'Стоимость не заполнена'

    else:
         print 'Документ должен содержать три рабочих листа'