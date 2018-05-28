'''
Created on March 27, 2018

@author: LIZHAOQ1
'''
#encoding=utf-8
import re
import xlrd
import xlsxwriter
import win32com.client as win32
import os
from vsslogging import logger


def readExcelGetWorkSheet(file_name):
    workbook = xlrd.open_workbook(file_name)
    sheets = workbook.sheet_names()
    worksheet = workbook.sheet_by_name(sheets[0])
    return worksheet

def getSheetId(worksheet):
    for i in range(1, worksheet.nrows):
        sheetid = worksheet.cell_value(i, 0)
        yield sheetid

    '''
    if worksheet.nrows > 1:
        for i in range(1, worksheet.nrows):
            sheetid = worksheet.cell_value(i, 0)
            yield sheetid
    else:
        print('无付款单！')
    '''

def requestSheetUrlGetSheet(vsObject,generator_sheetid):  #generator_sheetid = readExcelGetSheetId(file_name) 生成器无敌
    for sheet_id in generator_sheetid:
        sheet_url = 'http://vss.crv.com.cn/scm/DaemonMain?cmid=3020302000&operation=print&service=Paymentvoucher&sheetid=%s' % (
        sheet_id)
        response = vsObject.s.get(sheet_url, headers=vsObject.headers, timeout=30)
        colname = ['sheetid', 'taxername', 'venderid', 'rcvname',
                   'rcvbank', 'rcvaccno', 'paymodeid', 'paymodename',
                   'planpaydate', 'payamtToChinese', 'pnpayamt', 'pnsheetid',
                   'planpaydate', 'pnpayamt', 'payamt']
        value_list = []
        for name in colname:
            try:
                name = re.search("(?<=<%s>).*(?=</%s>)" % (name, name), response.text).group(0)
                value_list.append(name)
            except Exception as e:
                value_list.append(' ')
                print(e)
                logger.error(e)
        yield value_list, sheet_id


def writeExcel(username, generator_valuelist_sheetid, title):  #generrator_value_list = reuquestSheetUrlGetSheet(generrator_sheetid) 生成器
    for value_list, sheet_id in generator_valuelist_sheetid:
        #print(value_list, sheet_id)
        file_name = username + sheet_id + '.xls'
        file_path = os.path.join(r'.\tablefile', title, file_name)
        workbook = xlsxwriter.Workbook(file_path)
        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({'bold': True})
        worksheet.write('A1', '单据编号', bold)
        worksheet.write('B1', value_list[0])
        worksheet.write('C1', '结算主体', bold)
        worksheet.write('D1', value_list[1])

        worksheet.write('A2', '供应商编码', bold)
        worksheet.write('B2', value_list[2])
        worksheet.write('C2', '供应商名称', bold)
        worksheet.write('D2', value_list[3])

        worksheet.write('A3', '开户银行', bold)
        worksheet.write('B3', value_list[4])
        worksheet.write('C3', '银行账号', bold)
        worksheet.write('D3', value_list[5])

        worksheet.write('A4', '支付方式', bold)
        worksheet.write('B4', value_list[6] + ' ' + value_list[7])
        worksheet.write('C4', '计划付款日期', bold)
        worksheet.write('D4', value_list[8])

        worksheet.write('A5', '付款金额（大写）', bold)
        worksheet.write('B5', value_list[9])
        worksheet.write('C5', '付款金额（小写）', bold)
        worksheet.write('D5', value_list[10])

        worksheet.write('A6', '供应商盖章 ', bold)
        worksheet.write('B6:D6', '（盖章处）')

        worksheet.write('A8:D8', '对账单据', bold)


        worksheet.write('A9', '结算单号', bold)
        worksheet.write('B9','计划日期', bold)
        worksheet.write('C9', '结算单应付金额', bold)
        worksheet.write('D9', '本次付款金额', bold)
        worksheet.write('E9', '单据说明', bold)

        worksheet.write('A10', value_list[11])
        worksheet.write('B10', value_list[12])
        worksheet.write('C10', value_list[13])
        worksheet.write('D10', value_list[14])
        worksheet.write('E10', ' ')

        worksheet.write('A11:B11', '合计：')
        worksheet.write('C11', value_list[13])
        worksheet.write('D11', value_list[14])
        worksheet.write('E11', ' ')

        yield file_name



def sendEmail2(generator_excelName, title, receiver):    #generator_writeExcel =
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    receivers = receiver
    mail.To = receivers[0]
    mail.Subject = title + '\n'
    mail.Body = title
    for filename in generator_excelName:
        logger.debug(filename)
        try:
            mail.Attachments.Add(os.path.dirname(os.path.realpath(__file__)) + '\\tablefile' + '\\' + title + '\\' + filename)
        except Exception as e:
            mail.Body = filename + '发送失败！' + '\n' + str(e)  #e.value
            logger.error(e)
            continue
    mail.Send()
    print('付款单邮件发送成功！')
    logger.info(u'付款单邮件发送成功！')

    '''
    用for循环调用generator时，发现拿不到generator的return语句的返回值。如果想要拿到返回值，必须捕获StopIteration错误，返回值包含在StopIteration的value中
    while True:
        try:
        except StopIteration as e:
            print('Generator return value:', e.value)
            break
    '''





















