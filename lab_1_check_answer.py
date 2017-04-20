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

# wb2 = load_workbook('lab1_template.xlsx')
# ws = wb2[wb2.get_sheet_names()[0]]

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


def range_is_money_rub_format(ws, range):
    rows = ws[range]
    for row in rows:
        for cell in row:
            if not '₽' in cell.number_format:
                return {'status': False, 'message': 'Money rub format invalid'}
    return {'status': True, 'message': 'Money rub format valid'}

def range_is_money_dollar_format(ws, range):
    rows = ws[range]
    for row in rows:
        for cell in row:
            if not '$' in cell.number_format:
                return {'status': False, 'message': 'Money dollar format invalid'}
    return {'status': True, 'message': 'Money dollar format valid'}

def range_is_formula_format(ws, range):
    rows = ws[range]
    for row in rows:
        for cell in row:
            if str(cell.value)[0] != '=':
                return False
    return True

def range_is_date_format(ws, range):
    rows = ws[range]
    for row in rows:
        for cell in row:
            if not cell.is_date:
                return {'status': False, 'message': 'Dates invalid'}
    return {'status': True, 'message': 'Dates valid'}

def range_values_array(ws, range):
    arr = []
    rows = ws[range]
    for row in rows:
        for cell in row:
            arr.append(cell.value)
    return arr

def range_values_array_numeric(ws, range):
    arr = []
    rows = ws[range]
    for row in rows:
        for cell in row:
            arr.append(float(cell.value))
    return arr

def print_range(ws, range):
    rows = ws[range]
    for row in rows:
        for cell in row:
            print cell.coordinate, cell.value


def calculate_correct_values(ws, range, employees, dollar_rate):
    data = {}
    data["names"] = employees
    try:
        rows = ws[range]
        values = []
        for row in rows:
            for cell in row:
                values.append(float(cell.value))
        data["salary"] = {"values": values}

        premium = []
        total = []
        income_tax = []
        amount_granted = []
        amount_granted_dollar = []

        premium_formula = []
        total_formula = []
        income_tax_formula = []
        amount_granted_formula = []
        amount_granted_dollar_formula = []



        for index, v in enumerate(data["salary"]["values"]):

            premium_val = v * 0.2
            premium.append(premium_val)
            premium_formula.append('=0,2*E'+str(index+5))

            total_val = premium_val + v
            total.append(total_val)
            total_formula.append('=F'+str(index+5)+'+E'+str(index+5))

            income_tax_val = round(total_val*0.13, 2)
            income_tax.append(income_tax_val)
            income_tax_formula.append('=G'+str(index+5)+'*0,13')

            amount_granted_val = (total_val - income_tax_val)
            amount_granted.append(amount_granted_val)
            amount_granted_formula.append('=G'+str(index+5)+'-H'+str(index+5))

            amount_granted_dollar_val = round(amount_granted_val/dollar_rate, 2)
            amount_granted_dollar.append(amount_granted_dollar_val)
            amount_granted_dollar_formula.append('=I'+str(index+5)+'/$C$14')

        data["premium"] = {"values": premium, "formula": premium_formula}
        data["total"] = {"values": total, "formula": total_formula}
        data["total_sum"] = { "values": sum(total), "formula": '=СУММ(G5:G11)'}
        data["income_tax"] = {"values": income_tax, "formula": income_tax_formula}
        data["income_tax_sum"] = { "values": sum(income_tax), "formula": '=СУММ(H5:H11)'}
        data["amount_granted"] = {"values": amount_granted, "formula": amount_granted_formula}
        data["amount_granted_sum"] = {"values": sum(amount_granted), "formula": '=СУММ(I5:I11)'}
        data["amount_granted_dollar"] = {"values": amount_granted_dollar, "formula": amount_granted_dollar_formula}
        data["amount_granted_dollar_sum"] = {"values": sum(amount_granted_dollar), "formula": '=СУММ(J5:J11)'}

        data["avg_salary"] = {"values": round(reduce(lambda x, y: x + y, amount_granted) / len(amount_granted), 2), "formula": '=СРЗНАЧ(I5:I11)'}
        data["min_salary"] = {"values": round(min(amount_granted), 2), "formula": '=МАКС(I5:I11)'}
        data["max_salary"] = {"values": round(max(amount_granted), 2), "formula": '=МИН(I5:I11)'}
        return data
    except:
        return False


def get_range_data(ws, range, ws_data_only):
    try:
        data = {}
        formula =  range_values_array(ws, range)
        data["formula"] = formula
        values = range_values_array_numeric(ws_data_only, range)
        data["values"]= values
        return data
    except:
        return False

def get_students_formulas_values(ws, s, e, ws_data_only):
    data = {}
    data["names"] = range_values_array(ws, 'B'+str(s)+':'+'B'+str(e))

    salary =  range_values_array_numeric(ws, 'E'+str(s)+':'+'E'+str(e))
    data["salary"] = {"values": salary}

    premium =  range_values_array(ws, 'F'+str(s)+':'+'F'+str(e))
    data["premium"] = {}
    data["premium"]["formula"] = premium
    premium = range_values_array_numeric(ws_data_only, 'F' + str(s) + ':' + 'F' + str(e))
    data["premium"]["values"]= premium

    total = range_values_array(ws, 'G'+str(s)+':'+'G'+str(e))
    data["total"] = {}
    data["total_sum"] = {}
    data["total"]["formula"] =  total
    total = range_values_array_numeric(ws_data_only, 'G' + str(s) + ':' + 'G' + str(e))
    data["total"]["values"] = total
    data["total_sum"]["values"] = round(float(ws_data_only['G' + str(e + 1)].value), 2)
    data["total_sum"]["formula"] = ws['G' + str(e + 1)].value

    income_tax =  range_values_array(ws, 'H'+str(s)+':'+'H'+str(e))
    data["income_tax"] = {}
    data["income_tax_sum"] = {}
    data["income_tax"]["formula"] =  income_tax
    income_tax = range_values_array_numeric(ws_data_only, 'H' + str(s) + ':' + 'H' + str(e))
    data["income_tax"]["values"] = income_tax
    data["income_tax_sum"]["values"] = round(float(ws_data_only['H' + str(e + 1)].value), 2)
    data["income_tax_sum"]["formula"] = ws['H' + str(e + 1)].value

    amount_granted =  range_values_array(ws, 'I'+str(s)+':'+'I'+str(e))
    data["amount_granted"] = {}
    data["amount_granted_sum"] = {}
    data["amount_granted"]["formula"] = amount_granted
    amount_granted = range_values_array_numeric(ws_data_only, 'I' + str(s) + ':' + 'I' + str(e))
    data["amount_granted"]["values"] = amount_granted
    data["amount_granted_sum"]["values"] = round(float(ws_data_only['I' + str(e + 1)].value), 2)
    data["amount_granted_sum"]["formula"] = ws['I' + str(e + 1)].value

    amount_granted_dollar =  range_values_array(ws, 'J'+str(s)+':'+'J'+str(e))
    data["amount_granted_dollar"] = {}
    data["amount_granted_dollar_sum"] = {}
    data["amount_granted_dollar"]["formula"] =  amount_granted_dollar
    amount_granted_dollar = range_values_array_numeric(ws_data_only, 'J' + str(s) + ':' + 'J' + str(e))
    data["amount_granted_dollar"]["values"] = amount_granted_dollar
    data["amount_granted_dollar_sum"]["values"] = round(float(ws_data_only['J' + str(e + 1)].value), 2)
    data["amount_granted_dollar_sum"]["formula"] = ws['J' + str(e + 1)].value

    data["avg_salary"] = {}
    data["max_salary"] = {}
    data["min_salary"] = {}

    data["avg_salary"]["value"] = round(float(ws_data_only['C'+ str(e + 4)].value), 2)
    data["max_salary"]["value"] = round(float(ws_data_only['C'+ str(e + 5)].value), 2)
    data["min_salary"]["value"] = round(float(ws_data_only['C'+ str(e + 6)].value), 2)

    data["avg_salary"]["formula"] = ws['C'+ str(e + 4)].value
    data["max_salary"]["formula"] = ws['C'+ str(e + 5)].value
    data["min_salary"]["formula"] = ws['C'+ str(e + 6)].value

    return data

def check_names(correct_data, student_data):
    if correct_data["names"] == student_data["names"]:
        return {'status': True, 'message': 'Names valid'}
    else:
        return {'status': False, 'message': 'Names invalid'}

def check_formats(student_ws):
    # Проверяем правильность форматирования дат поступления
    print "Формат дат поступления: ", range_is_date_format(student_ws, 'D5:D11')['status']

    # Проверяем правильность форматирования оклада
    print "Формат оклада: ", range_is_money_rub_format(student_ws, 'E5:E11')['status']

    # Проверяем правильность форматирования премии
    print "Формат премии: ", range_is_money_rub_format(student_ws, 'F5:F11')['status']

    # Проверяем правильность форматирования итого
    print "Формат итого: ", range_is_money_rub_format(student_ws, 'G5:G11')['status']

    # Проверяем правильность форматирования подоходного налога
    print "Формат подоходного налога: ", range_is_money_rub_format(student_ws, 'H5:H11')['status']

    # Проверяем правильность форматирования суммы к выдаче
    print "Формат суммы к выдаче: ", range_is_money_rub_format(student_ws, 'I5:I11')['status']

    # Проверяем правильность форматирования суммы к выдаче в долларах
    print "Формат суммы к выдаче в долларах: ", range_is_money_dollar_format(student_ws, 'J5:J11')['status']


def formulas_arrays_is_equal(f1, f2):
    clean_f1 = []
    clean_f2 = []
    for f in f1:
        clean_f1.append(f.replace(" ", "").lower().replace(",", "."))
    for f in f2:
        clean_f2.append(f.replace(" ", "").lower().replace(",", "."))

    return clean_f1 == clean_f2


def check_formulas(ws, ws_read_only, correct_data):
    premium = get_range_data(ws, 'F5:F11', ws_read_only)

    premium_formulas_msg = "Премии посчитаны: "
    # range_is_formula_format(ws, 'F5:F11')
    if premium and formulas_arrays_is_equal(premium["formula"], correct_data["premium"]["formula"]) and premium["values"] == correct_data["premium"]["values"]:
        premium_formulas_msg += 'True'
    else:
        premium_formulas_msg += 'False'

    print premium_formulas_msg

def ws_have_rule(ws, cells_range, operator, formula_value):
    for rule in ws.conditional_formatting.cf_rules.items():
        if cells_range == rule[0]:
            if operator == rule[1][0].operator:
                if int(rule[1][0].formula[0]) == formula_value:
                    return True
    return False


def check_answer(student_wb, student_wb_data_only, employees, dollar_rate):
    student_ws = student_wb[student_wb.get_sheet_names()[0]]
    ws_read_only = student_wb_data_only[student_wb_data_only.get_sheet_names()[0]]

    start_row = 5
    end_row = start_row + len(employees) - 1
    range_salary = 'E5:E11'

    correct_values_data = {}
    correct_values_data = calculate_correct_values(student_ws, range_salary, sorted(employees), dollar_rate)

    if correct_values_data:
        # Проверяем правильность ФИО (в т.ч. сортировку)
        # check_formats(student_ws)
        check_formulas(student_ws, ws_read_only, correct_values_data)

        # print ws_have_rule(student_ws, 'I5:I11', 'lessThan', 5000)

    else:
        message = 'Неверно заполнен столбец "Оклад"'

    student_values_data = {}
    #student_values_data = get_students_formulas_values(student_ws, start_row, end_row, ws_read_only)
    # print student_values_data
