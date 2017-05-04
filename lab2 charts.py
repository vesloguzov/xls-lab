# -*- coding: UTF-8 -*-
import sys
import random
import datetime

from openpyxl import Workbook, load_workbook
from openpyxl import Workbook, load_workbook
from openpyxl.styles import *
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.reader import reader
from openpyxl.chart.layout import Layout, ManualLayout
from zipfile import ZipFile, ZIP_DEFLATED
from lxml.etree import fromstring, tostring
from openpyxl.chart import (
    ScatterChart,
    Reference,
    Series,
)


reload(sys)
sys.setdefaultencoding('utf8')


def check_scatter_graphic(filename, data_x, data_y, num):
    analyze = {}
    analyze["errors"] = []
    analyze["chart_title"] = {}
    analyze["data_x"] = {}
    analyze["data_y"] = {}
    analyze["title_y"] = {}

    sourceFile = ZipFile(filename, 'r')
    charts = []; [charts.append(sourceFile.read(ch)) for ch in sourceFile.namelist() if 'charts/chart' in ch]
    charts_objects = []
    for chart in charts:
        try:
            clean_chart = fromstring(chart)
            for bad in clean_chart.xpath(".//c:extLst", namespaces={'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}):
                bad.getparent().remove(bad)
            clean_chart = tostring(clean_chart)
            l = reader(clean_chart)

            charts_objects.append(l)
        except:
            analyze["errors"].append('График не обнаружен!')
            return analyze

    obj = charts_objects[num]

    for chart in charts_objects:
        if chart.tagname == 'scatterChart':
            obj = chart
            break

    if obj == {}:
        analyze["errors"].append('График не обнаружен!')
    else:
        if obj.title != None:
            analyze["chart_title"]["message"] = 'Имя графика присвоено'
            analyze["chart_title"]["status"] = True
        else:
            analyze["chart_title"]["message"] = 'Имя графика не присвоено'
            analyze["chart_title"]["status"] = False
        try:
            for s in obj.ser:
                if data_x in s.xVal.numRef.f.replace(" ", ""):
                    analyze["data_x"]["message"] = 'Данные для оси x выбраны верно'
                    analyze["data_x"]["status"] = True
                else:
                    analyze["data_x"]["message"] = 'Данные для оси x выбраны неверно'
                    analyze["data_x"]["status"] = False
        except:
            analyze["data_x"]["message"] = 'Данные для оси x выбраны неверно'
            analyze["data_x"]["status"] = False

        try:
            for s in obj.ser:
                if data_y in s.yVal.numRef.f.replace(" ", ""):
                    analyze["data_y"]["message"] = 'Данные для оси y выбраны верно'
                    analyze["data_y"]["status"] = True
                else:
                    analyze["data_y"]["message"] = 'Данные для оси y выбраны неверно'
                    analyze["data_y"]["status"] = False
        except:
            analyze["data_y"]["message"] = 'Данные для оси y выбраны неверно'
            analyze["data_y"]["status"] = False

        try:
            if obj.y_axis.title != None:
                analyze["title_y"]["message"] = 'Подпись осей выполнена'
                analyze["title_y"]["status"] = True
            else:
                analyze["title_y"]["message"] = 'Подпись осей не выполнена'
                analyze["title_y"]["status"] = False
        except:
            analyze["title_y"]["message"] = 'Подпись осей не выполнена'
            analyze["title_y"]["status"] = False

    return  analyze

print check_scatter_graphic('lab2_correct.xlsx', '$B$5:$B$11', '$I$5:$I$11', 0)

# Лист 1
# $A$4:$A$28 #Ось x
# $B$4:$B$28 #Ось y

# Лист 2
# $A$4:$A$34 #Ось x
# $B$4:$B$34 #Ось y


# response_message = {}
# response_message["ws1"]["data"] = {}
# response_message["ws1"]["graphic"] = {}

# Название графика
# print obj.title

# SCATTER CHART

# print obj.y_axis.title
# for s in obj.ser:
#    print "Диапазон данных по оси x: ", s.xVal.numRef.f
#    print "Диапазон данных по оси y: ", s.yVal.numRef.f