#encoding=utf-8
from vssspider import VssSpider
import time
import re
import os
import xlsxwriter
from bs4 import BeautifulSoup
from vsslogging import logger
import getDate

class VssSpiderPD(VssSpider):

    def getUrlRowXmlList(self, url):
        for i in range(0, 20):
            try:
                r = self.s.get(url, headers=self.headers, timeout=60)
                logger.debug(u'促销明细' + str(r.status_code))
                break
            except Exception as e:
                time.sleep(2)
                print(e)
                logger.error(url + str(e))
                continue

        soup = BeautifulSoup(r.text, "xml")
        rowxmllist = soup.find_all('Row')
        return rowxmllist


    def getRowValueList(self, rowxmllist):
        iter_rowxmllist = iter(rowxmllist)
        for rowxml in iter_rowxmllist:
            #print(str(rowxml))
            '''
            r'<ss:Cell><ss:Data ss:Type="String">1006031110</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="String">美太芭比（上海）贸易有限公司</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="String">JV华东</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="String">204368</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="String">汾湖店</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="Number">2562240</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="String">6947731024773</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="String">CARS赛车总动员变形麦克DVF39</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="Number">1</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="Number">70.28</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="Number">82.23</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="String">非促销</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="String"/></ss:Cell><ss:Cell><ss:Data ss:Type="String"/></ss:Cell><ss:Cell><ss:Data ss:Type="String">-</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="Number">49191411</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="String">其他玩具</ss:Data></ss:Cell><ss:Cell><ss:Data ss:Type="String">2018-03-24 00:00:00</ss:Data></ss:Cell></ss:Row>'
            <ss:Cell><ss:Data ss:Type="String"/></ss:Cell>  注意这种情况re匹配
            [('1006031110', '</ss:Data>'), ('美太芭比（上海）贸易有限公司', '</ss:Data>'), ('JV华东', '</ss:Data>'), ('204368', '</ss:Data>'), ('汾湖店', '</ss:Data>'), ('6947731024773', '</ss:Data>'), ('CARS赛车总动员变形麦克DVF39', '</ss:Data>'), ('非促销', '</ss:Data>'), ('', ''), ('', ''), ('-', '</ss:Data>'), ('其他玩具', '</ss:Data>'), ('2018-03-24 00:00:00', '</ss:Data>')]
            处理上面这个list  用zip(*)  zip逆操作
            '''
            #rowvaluelist = re.findall(r'<ss:Data ss:Type="[A-Za-z]+">(.*?)</ss:Data>', str(rowxml))
            rowvaluelist = re.findall(r'<ss:Data ss:Type="[A-Za-z]+"/?>(.*?)(</ss:Data>)?</ss:Cell>', str(rowxml))
            rowvaluetuple = list(zip(*rowvaluelist))[0]   #zip(*)  逆操作
            #print(rowvaluelist)
            yield rowvaluetuple

    def writeExcel(self, gene_rowvaluetuple, user_name, title):
        localtime = getDate.getLocalTime()
        filename = user_name + title + localtime + '.xlsx'
        relativepath = r'.\tablefile'   #相对路径
        filepath = os.path.join(relativepath, title, filename)  #组合相对路径

        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet(title)
        i = 1
        for rowvaluelist in gene_rowvaluetuple:
            #print(rowvaluelist)
            worksheet.write_row('A{}'.format(i), rowvaluelist)
            i+=1

        workbook.close()
        print(filename + " 导出完成！")
        logger.info(filename + " 导出完成！")
        return filename, filepath





