'''
Created on Mar 20, 2018

@author: lizhaoq1
'''
import getDate
from vssspider import VssSpider
from vsslogging import logger
import handleExcel
import os
import getDate


def theLastInventory(user_list, receiver):
    title = '最新库存'
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
        url = r'http://vss.crv.com.cn/scm/DaemonSCMDownloadReport?reportname=cmexcel&class=CurrtentInventory4Scm&cmid=9000003000&buid=%s&venderid=%s&' % (buid, venderid)
        v = VssSpider(username, password)
        txtcheck = v.getCheckCode()
        v.longonVender(txtcheck)

        is_exist, file_name, file_path = v.exportExcel(url, username, title)
        if is_exist:
            filepath_list.append(file_path)
            filewarn[username] = "导出成功！"
        else:
            filewarn[username] = "导出失败！"
        #v.sendEmail(file_name, title, receiver)   #每个账号表都发送
    warntitle = str(filewarn)
    logger.info(warntitle)
    file_absolute_path = handleExcel.handel_main(filepath_list, title)
    handleExcel.send_email_total(title, file_absolute_path, receiver, warntitle)

    print('最新库存完成')
    logger.info(u'最新库存完成')
#theLastInventory()