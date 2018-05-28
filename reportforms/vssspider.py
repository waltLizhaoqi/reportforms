'''
Created on Mar 20, 2018

@author: lizhaoq1
'''
#encoding=utf-8
import requests
import os
import time
import random
import getDate
import re
from bs4 import BeautifulSoup
from PIL import Image
import pytesseract
import win32com.client as win32
from contextlib import closing
from retry import retry
from vsslogging import logger

class VssSpider(object):
    def __init__(self, name, password):
        self.logon_url = 'http://vss.crv.com.cn/scm/logon/logon.jsp'
        self.code_url = 'http://vss.crv.com.cn/scm/DaemonCode'
        self.logon_vender = 'http://vss.crv.com.cn/scm/DaemonLogonVender?'
        self.main_url = 'http://vss.crv.com.cn/scm/logon/main.htm'
        self.headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'en-US,en;q=0.9',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            # 'Cookie':'route=fe207c8a23a8c55f8a14b32a664aace4; JSESSIONID=9D445BD319014703580EEBDC5E9089C9.wjvsmvpca01vsmscm01; IESESSION=alive; _qddaz=QD.bemviv.giswno.jdjyd9ft; pgv_pvi=6044569600; pgv_si=s753846272; _qddab=4-593fzw.jdl9z7gg; route=7d8fe1888a882023791dcb0f8d617654',
            'Host': 'vss.crv.com.cn',
            'Referer': 'http://vss.crv.com.cn/scm/logon/logon.jsp',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36',
        }
        self.user_name = name
        self.user_password = password
        self.s = requests.session()   #会话关闭

    def getCheckCode(self):
        # 验证码连接 from PIL import Image  import pytesseract
        while True:
            try:
                response = self.s.get(self.code_url, headers=self.headers)
            except Exception as e:
                time.sleep(2)
                print(e)
                logger.error(e)
                continue
            break
        picture = response.content
        # 将验证码写入本地
        local = open("checkcode.jpg", "wb")
        local.write(picture)
        local.close()

        im = Image.open('checkcode.jpg')
        # 把彩色图像转化为灰度图像。RBG转化到HSI彩色空间，采用L分量
        imgry = im.convert('L')

        # out = imgry.point(lambda x:0 if x<140 else 255, '1')
        # 二值化处理
        threshold = 140
        table = []
        for i in range(256):
            if i < threshold:
                table.append(0)
            else:
                table.append(1)
        out = imgry.point(table, '1')
        #out.show()

        out.save('checkcode2.jpg')
        im2 = Image.open('checkcode2.jpg')
        resizedIm = im2.resize((600, 200))
        resizedIm.save('checkcode3.jpg')

        #resizedIm.show()
        #pytesseract.pytesseract.tesseract_cmd = os.path.dirname(os.path.realpath(__file__)) + '\\' +  'Tesseract-OCR\\tesseract.exe'   #自动获取路径
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'  #tesseract的路径 3种方式 还有一种 pytesseract.py文件里 path环境变量
        text_check = pytesseract.image_to_string(resizedIm).replace(' ', '')
        os.remove('checkcode.jpg')
        os.remove('checkcode2.jpg')
        os.remove('checkcode3.jpg')

        #print(text_check)
        return text_check

    def longonVender(self, textcheck):
        logon_get_data = {
            'site': '0',
            'action': 'logon',
            'logonid': self.user_name,
            'password': self.user_password,
            'checkcode': textcheck,
        }
        for i in range(0,20):
            try:
                r = self.s.get(self.logon_vender, params=logon_get_data, headers=self.headers, timeout=60)
                logger.debug(self.logon_vender + str(r.status_code))
                break

                #print(r.text)
                # print(re.findall(r"<note>验证码错误!</note>", r.text))
            except Exception as e:  #requests.exceptions.ReadTimeout
                time.sleep(2)
                print(e)
                continue

        try:
            if re.findall(r"<note>验证码错误!</note>", r.text):
                time.sleep(random.randint(1, 5))
                text_check = self.getCheckCode()
                self.longonVender(text_check)
        except Exception as e:
            print(e)
            print('网络请求失败，请稍后再试！')
            logger.critical(e)
            logger.critical(u'网络请求失败，请稍后再试！')

    # @retry(tries=5, delay=2)    #隔两秒 失败重试
    def exportExcel(self, url, user_name, title):
        localtime = getDate.getLocalTime()
        filename = user_name + title + localtime + '.xls'
        relativepath = r'.\tablefile'   #相对路径
        filepath = os.path.join(relativepath, title, filename)  #组合相对路径
        isexist = False

        for i in range(0,20):
            try:
                with closing(self.s.get(url, headers=self.headers, stream=True, allow_redirects=False, timeout=60)) as response:  #请求超时,登录超时情况处理
                    logger.debug(response.status_code)
                    if re.search(r"<code>-1</code>", response.text) is None:
                        logger.info(user_name + title + "访问成功！")
                        with open(filepath, 'wb') as file: #filename
                            for chunk in response.iter_content(chunk_size=1024):  #分块下载
                                if chunk:
                                    file.write(chunk)

                            isexist = True
                            print(filename + '导出完成！')
                            logger.info(filename + u'导出完成！')
                        break
                    else:
                        continue
            except Exception as e:
                logger.warn(user_name + title + "访问失败！" + str(e))
                continue

        logger.debug(user_name + title + "访问次数：" + str(i))
        return isexist, filename, filepath


    def sendEmail(self, filename, title, receiver):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        receivers = receiver
        mail.To = receivers[0]
        mail.Subject = title + '\n'
        mail.Body = title
        try:
            mail.Attachments.Add(os.path.dirname(os.path.realpath(__file__)) + '\\tablefile' + '\\' + title + '\\' + filename)  #获取本文件路径 os.path.dirname(os.path.realpath(__file__)) + '\\' + filename
        except Exception as e:
            mail.Body = filename + '发送失败！' + '\n' + str(e)  # e.value
            logger.error(e)
        mail.Send()
        print(filename + '邮件发送成功！')
        logger.info(filename + u'邮件发送成功！')



