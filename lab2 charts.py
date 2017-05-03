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

    obj = charts_objects[0]
    print obj.title
    #obj.y_axis.title.tx.rich.p

    # obj.x_axis.title = 'SizeXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    # obj.y_axis.title = 'SizeYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY'

    # print obj

    # for s in obj.ser:
    #    print "Диапазон данных по оси x: ", s.xVal.numRef.f
    #    print "Диапазон данных по оси y: ", s.yVal.numRef.f


check_bar_graphic('lab2_correct.xlsx', '$B$5:$B$11', '$I$5:$I$11')

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