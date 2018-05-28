'''
Created on Mar 20, 2018

@author: lizhaoq1
'''
import getDate
from vssspider import VssSpider
from vsslogging import logger
import handleExcel
from vssspiderPD import VssSpiderPD

min_date, max_date = getDate.getBefore1Month()
def promotionDetails(user_list, receiver):
    title = '促销明细'
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
        url = r'http://vss.crv.com.cn/scm/DaemonSCMDownloadReport?reportname=prom&buid=%s&venderid=%s&sdate_min=%s&sdate_max=%s' % (buid, venderid, min_date, max_date)
        #v = VssSpider(username, password)
        #txtcheck = v.getCheckCode()
        #v.longonVender(txtcheck)

        #file_name, file_path = v.exportExcel(url, username, title)
        #v.sendEmail(file_name, title, receiver)
        #filepath_list.append(file_path)
        vpd = VssSpiderPD(username, password)
        txtcheck = vpd.getCheckCode()
        vpd.longonVender(txtcheck)

        rowxmllist = vpd.getUrlRowXmlList(url)
        gene_rowvaluelist = vpd.getRowValueList(rowxmllist)
        file_name, file_path = vpd.writeExcel(gene_rowvaluelist, username, title)
        filepath_list.append(file_path)
        filewarn[username] = "导出成功！"

    warntitle = str(filewarn)
    logger.info(warntitle)
    file_absolute_path = handleExcel.handel_main(filepath_list, title)
    handleExcel.send_email_total(title, file_absolute_path, receiver, warntitle)
    print('促销明细完成')
    logger.info(u'促销明细完成')

#promotionDetails()