'''
Created on Mar 20, 2018

@author: lizhaoq1
'''
# -*- coding: utf-8 -*-
import getDate
import logging.handlers
from vssspider import VssSpider
import vssspiderPO
from vsslogging import logger
import handleExcel

# min_date, max_date = getDate.getBefore1Month()
min_date, max_date = getDate.getBefore60Day()
def expensesDocuments(user_list, receiver):
    title = '费用单据'
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
        url = r'http://vss.crv.com.cn/scm/DaemonMainDownload?operation=cmexcel&cmid=3020304000&service=ChargeSum&buid=%s&venderid=%s&docdate_min=%s&docdate_max=%s&' % (buid, venderid, min_date, max_date)
        v = VssSpider(username, password)
        txtcheck = v.getCheckCode()
        v.longonVender(txtcheck)

        is_exist, file_name, file_path = v.exportExcel(url, username, title)
        '''
        work_sheet = vssspiderPO.readExcelGetWorkSheet(file_name)
        if work_sheet.nrows > 1:
            v.sendEmail(file_name, title, receiver)
        else:
            print(username + min_date + '-' + max_date + "无费用单据！")
            logger.warn(username + min_date + '-' + max_date + u'无费用单据！')
        '''
        if is_exist:
            filepath_list.append(file_path)
            filewarn[username] = "导出成功！"
        else:
            filewarn[username] = "导出失败！"

    warntitle = str(filewarn)
    logger.info(warntitle)
    file_absolute_path = handleExcel.handel_main(filepath_list, title)
    handleExcel.send_email_total(title, file_absolute_path, receiver, warntitle)
    print('费用单据完成')
    logger.info(u'费用单据完成')
#expensesDocuments()