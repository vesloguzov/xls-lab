import random
import datetime
import time

def strTimeProp(start, end, format, prop):
    stime = time.mktime(time.strptime(start, format))
    etime = time.mktime(time.strptime(end, format))
    ptime = stime + prop * (etime - stime)
    return time.strftime(format, time.localtime(ptime))

startDate = "01.01.2000"
endDate = "01.01.2018"

def randomDate(start=startDate, end=endDate):
    prop = random.random()
    return strTimeProp(start, end, '%d.%m.%Y', prop)
