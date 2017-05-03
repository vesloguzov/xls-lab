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

def check_bar_graphic(filename, data_x, data_y):
    analyze = {}
    analyze["bar_chart"] = {}
    analyze["bar_chart"]["errors"] = []
    analyze["bar_chart"]["data_x"] = {}
    analyze["bar_chart"]["data_y"] = {}
    analyze["bar_chart"]["title_x"] = {}
    analyze["bar_chart"]["title_y"] = {}
    analyze["bar_chart"]["chart_title"] = {}

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
            analyze["bar_chart"]["errors"].append('График не обнаружен!')
            return analyze
    obj = {}
    for chart in charts_objects:
        if chart.tagname == 'barChart':
            obj = chart
            break

    if obj == {}:
        analyze["errors"].append('Тип графика выбран неверно')
    else:
        if obj.title != None:
            analyze["bar_chart"]["chart_title"]["message"] = 'Имя графика присвоено'
            analyze["bar_chart"]["chart_title"]["status"] = True
        else:
            analyze["bar_chart"]["chart_title"]["message"] = 'Имя графика не присвоено'
            analyze["bar_chart"]["chart_title"]["status"] = False

        try:
            for s in obj.ser:
                if data_x in s.cat.strRef.f.replace(" ", ""):
                    analyze["bar_chart"]["data_x"]["message"] = 'Данные для оси x выбраны верно'
                    analyze["bar_chart"]["data_x"]["status"] = True
                else:
                    analyze["bar_chart"]["data_x"]["message"] = 'Данные для оси x выбраны неверно'
                    analyze["bar_chart"]["data_x"]["status"] = False

        except:
            analyze["bar_chart"]["data_x"]["message"] = 'Данные для оси x выбраны неверно'
            analyze["bar_chart"]["data_x"]["status"] = False
        try:
            for s in obj.ser:
                if data_y in s.val.numRef.f.replace(" ", ""):
                    analyze["bar_chart"]["data_y"]["message"] = 'Данные для оси y выбраны верно'
                    analyze["bar_chart"]["data_y"]["status"] = True
                else:
                    analyze["bar_chart"]["data_y"]["message"] = 'Данные для оси y выбраны неверно'
                    analyze["bar_chart"]["data_y"]["status"] = False
        except:
            analyze["bar_chart"]["data_y"]["message"] = 'Данные для оси y выбраны неверно'
            analyze["bar_chart"]["data_y"]["status"] = False

        try:
            if obj.x_axis.title != None:
                analyze["bar_chart"]["title_x"]["message"] = 'Подпись оси x выполнена'
                analyze["bar_chart"]["title_x"]["status"] = True
            else:
                analyze["bar_chart"]["title_x"]["message"] = 'Подпись оси x не выполнена'
                analyze["bar_chart"]["title_x"]["status"] = False
        except:
            analyze["bar_chart"]["title_x"]["message"] = 'Подпись оси x не выполнена'
            analyze["bar_chart"]["title_x"]["status"] = False

        try:
            if obj.y_axis.title != None:
                analyze["bar_chart"]["title_y"]["message"] = 'Подпись оси y выполнена'
                analyze["bar_chart"]["title_y"]["status"] = True
            else:
                analyze["bar_chart"]["title_y"]["message"] = 'Подпись оси y не выполнена'
                analyze["bar_chart"]["title_y"]["status"] = False
        except:
            analyze["bar_chart"]["title_y"]["message"] = 'Подпись оси y не выполнена'
            analyze["bar_chart"]["title_y"]["status"] = False

    return analyze


print check_bar_graphic('8BD8D6k.xlsx', '$B$5:$B$11', '$I$5:$I$11')


#
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