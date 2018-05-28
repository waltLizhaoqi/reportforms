'''
Created on Mar 20, 2018

@author: lizhaoq1
'''
#encoding=utf-8
import calendar
import time
import datetime
def getLocalTime():
    localtime = time.strftime("%Y-%m-%dT%H-%M-%S", time.localtime())
    return localtime

def getThisMonth():
    day_now = time.localtime()

    day_begin = '%d-%02d-01' % (day_now.tm_year, day_now.tm_mon)
    day_end = time.strftime("%Y-%m-%d", day_now)

    return day_begin,day_end

def getBefore1Month():
    day_now = time.localtime()

    if day_now.tm_mon > 1:
        needmonth = day_now.tm_mon - 1
        day_begin = '%d-%02d-01' % (day_now.tm_year, needmonth)  # 月初肯定是1号
        wday, monthRange = calendar.monthrange(day_now.tm_year, needmonth)  # 得到本月的天数 第一返回为月第一日为星期几（0-6）, 第二返回为此月天数
        day_end = '%d-%02d-%02d' % (day_now.tm_year, needmonth, monthRange)
    elif day_now.tm_mon < 2:
        needyear = day_now.tm_year - 1
        day_begin = '%d-%02d-01' % (needyear, 12)
        wday, monthRange = calendar.monthrange(needyear, 12)
        day_end = '%d-%02d-%02d' % (needyear, 12, monthRange)


    return day_begin, day_end

def getBefore3Month():
    day_now = time.localtime()

    if day_now.tm_mon > 3:
        day_begin = '%d-%02d-01' % (day_now.tm_year, day_now.tm_mon - 3)
        wday, monthRange = calendar.monthrange(day_now.tm_year, day_now.tm_mon - 1)
        day_end = '%d-%02d-%02d' % (day_now.tm_year, day_now.tm_mon - 1, monthRange)
    elif day_now.tm_mon < 4:
        needmonth = 9 + day_now.tm_mon
        day_begin = '%d-%02d-01' % (day_now.tm_year - 1, needmonth)
        if day_now.tm_mon < 2:
            wday, monthRange = calendar.monthrange(day_now.tm_year - 1, 12)
            day_end = '%d-%02d-%02d' % (day_now.tm_year - 1, 12, monthRange)
        elif day_now.tm_mon > 1:
            wday, monthRange = calendar.monthrange(day_now.tm_year, day_now.tm_mon - 1)
            day_end = '%d-%02d-%02d' % (day_now.tm_year, day_now.tm_mon - 1, monthRange)

    return day_begin, day_end

def getBefore90Day():
    now_time = datetime.datetime.now()
    str_now_time = now_time.strftime('%Y-%m-%d')

    end_time = now_time + datetime.timedelta(days=-1)
    str_end_time = end_time.strftime('%Y-%m-%d')

    begin_time = now_time + datetime.timedelta(days=-90)
    str_begin_time = begin_time.strftime('%Y-%m-%d')
    return str_begin_time, str_end_time

def getBefore7Day():
    now_time = datetime.datetime.now()
    str_now_time = now_time.strftime('%Y-%m-%d')

    end_time = now_time + datetime.timedelta(days=-1)
    str_end_time = end_time.strftime('%Y-%m-%d')

    begin_time = now_time + datetime.timedelta(days=-7)
    str_begin_time = begin_time.strftime('%Y-%m-%d')

    return str_begin_time, str_end_time

def getBefore60Day():
    now_time = datetime.datetime.now()
    str_now_time = now_time.strftime('%Y-%m-%d')


    begin_time = now_time + datetime.timedelta(days=-59)
    str_begin_time = begin_time.strftime('%Y-%m-%d')
    return str_begin_time, str_now_time



