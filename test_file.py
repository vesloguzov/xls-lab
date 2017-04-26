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

sourceFile = ZipFile('lab2_correct.xlsx', 'r')
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
        print 'Ошибка разбора графика!'


obj = charts_objects[0]

print obj.title


#     chart.title = cs.chart.title
#     chart.layout = plot.layout
#     chart.legend = cs.chart.legend

# self.legendPos = legendPos
# self.legendEntry = legendEntry
# self.layout = layout
# self.overlay = overlay
# self.spPr = spPr
# self.txPr = txPr

# print obj.title.text.rich
# for p in l.y_axis.title.tx:
# for p in obj.y_axis.title.tx.rich.p:
#     print p.r
#
# print 'KEK'
# # print l.x_axis
#

for s in obj.ser:
    # print s.xVal.numRef.numCache # Значения числовые

    print "Диапазон данных по оси x: ", s.xVal.numRef.f # Диапазон данных
    print "Диапазон данных по оси y: ", s.yVal.numRef.f