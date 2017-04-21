# -*- coding: UTF-8 -*-
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.reader import reader
from openpyxl.chart.layout import Layout, ManualLayout

reload(sys)
sys.setdefaultencoding('utf8')


def check_answer(student_wb, student_wb_data_only, data):
    try:
        student_ws_1 = student_wb[student_wb.get_sheet_names()[0]]
        student_ws_2 = student_wb[student_wb.get_sheet_names()[1]]
        student_ws_3 = student_wb[student_wb.get_sheet_names()[2]]
        ws_read_only_1 = student_wb_data_only[student_wb_data_only.get_sheet_names()[0]]
        ws_read_only_2 = student_wb_data_only[student_wb_data_only.get_sheet_names()[1]]
        ws_read_only_3 = student_wb_data_only[student_wb_data_only.get_sheet_names()[2]]

        print data
    except:
        print 'Документ должен содержать три рабочих листа'