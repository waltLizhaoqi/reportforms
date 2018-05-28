'''
Created on Mar 20, 2018

@author: lizhaoq1
'''
# -*- coding: utf-8 -*-
import getDate
from vssspider import VssSpider
import vssspiderPO
import handleExcel
from vsslogging import logger
min_date, max_date = getDate.getBefore1Month()


def paymentOrder(user_list, receiver):
    title = '付款单'
    '''
    receiver = ['zhaoqi.li@mattel.com;']
    user_list = [('161006031101','123456','16','1006031101'), ('18V08747','123456','18','V08747'),
                ('1558098','123456','15','58098'), ('131006031102','123456','13','1006031102'),
                ('121006031103','123456','12','1006031103'), ('881006031105','123456','88','1006031105'),
                ('321006031107','123456','32','1006031107'), ('311006031110','123456','31','1006031110')]
    '''
    filewarn = dict()
    filepath_list = []
    for username, password, buid, venderid in user_list:
        #url = r'http://vss.crv.com.cn/scm/DaemonMainDownload?operation=cmexcel&cmid=3020302000&service=Paymentvoucher&buid=%s&venderid=%s&' % (buid, venderid)
        # url = r'http://vss.crv.com.cn/scm/DaemonMainDownload?operation=cmexcel&cmid=3020302000&service=Paymentvoucher&buid=%s&venderid=%s&date_min=%s&date_max=%s' % (buid, venderid, min_date, max_date)
        url = r'http://vss.crv.com.cn/scm/DaemonMainDownload?operation=cmexcel&cmid=3020302000&service=Paymentvoucher&buid=%s&venderid=%s&' % (buid, venderid)
        vsObject = VssSpider(username, password)
        txtcheck = vsObject.getCheckCode()
        vsObject.longonVender(txtcheck)

        is_exist, file_name, file_path = vsObject.exportExcel(url, username, title)
        # work_sheet = vssspiderPO.readExcelGetWorkSheet(file_path)
        # if work_sheet.nrows > 1:
        #     vsObject.sendEmail(file_name, title, receiver)
        #     generator_sheetid = vssspiderPO.getSheetId(work_sheet)
        #     generator_valuelist_sheetid = vssspiderPO.requestSheetUrlGetSheet(vsObject, generator_sheetid)
        #     generator_excelName = vssspiderPO.writeExcel(username, generator_valuelist_sheetid, title)
        #     vssspiderPO.sendEmail2(generator_excelName, title, receiver)
        # else:
        #     print("无付款单！")
        #     logger.info(u"无付款单！")
        if is_exist:
            filepath_list.append(file_path)
            filewarn[username] = "导出成功！"
        else:
            filewarn[username] = "导出失败！"

        '''
        generator_sheetid = vssspiderPO.readExcelGetSheetId(file_name)
        generator_valuelist_sheetid = vssspiderPO.requestSheetUrlGetSheet(vsObject, generator_sheetid)
        generator_excelName = vssspiderPO.writeExcel(username, generator_valuelist_sheetid)
        vssspiderPO.sendEmail2(generator_excelName, title, receiver)
        '''
    warntitle = str(filewarn)
    logger.info(warntitle)
    file_absolute_path = handleExcel.handel_main(filepath_list, title)
    handleExcel.send_email_total(title, file_absolute_path, receiver, warntitle)
    print('付款单完成')
    logger.info(u'付款单完成')

#paymentOrder()