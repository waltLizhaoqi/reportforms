'''
Created on Mar 20, 2018

@author: lizhaoq1
'''
#encoding=utf-8
from vssspider import VssSpider
import getDate
import handleExcel
from vsslogging import logger

begintime, endtime = getDate.getBefore7Day()
def salesOfDailyReport(user_list, receiver):
    title = '销售日报'
    '''
    receiver = ['zhaoqi.li@mattel.com;']
    user_list = [('161006031101','123456','16'),('18V08747','123456','18'),
                ('1558098','123456','15'),('131006031102','123456','13'),
                ('121006031103','123456','12'),('881006031105','123456','88'),
                ('321006031107','123456','32'),('311006031110','123456','31')]   #经过分析url中有一个参数 buid= 用户名的前两位
    '''
    print(begintime + '-' + endtime + '数据导出中...')
    filewarn = dict()
    filepath_list = []
    for username, password, buid in user_list:
        v = VssSpider(username, password)
        txtcheck = v.getCheckCode()
        v.longonVender(txtcheck)
        url = 'http://vss.crv.com.cn/scm/DaemonSCMDownloadReport?reportname=cmexcel&class=SaleDaily4Scm&cmid=9000002000&buid=%s&sdate_min=%s&sdate_max=%s&' % (buid, begintime, endtime)
        #file_name = v.exportExcel(url, username, title)
        #v.sendEmail(file_name, title, receiver)
        #time.sleep(random.randint(1, 5))
        is_exist, file_name, file_path = v.exportExcel(url, username, title)
        if is_exist:
            filepath_list.append(file_path)
            filewarn[username] = "导出成功！"
        else:
            filewarn[username] = "导出失败！"
    warntitle = str(filewarn)
    logger.info(warntitle)
    file_absolute_path = handleExcel.handel_main(filepath_list, title)
    handleExcel.send_email_total(title, file_absolute_path, receiver, warntitle)
    print('销售日报完成')

#日期是今天的前一天（昨天）往前数7天
