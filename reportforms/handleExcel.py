# -*- coding:utf-8 -*-
import pandas as pd
from vsslogging import logger
import getDate
import os
import win32com.client as win32


def merge_excel(filepathlist):
    """合并数据"""
    all_data = pd.DataFrame()
    for file in filepathlist:   #遍历path路径下的所有文件名
        logger.debug(file)
        df = pd.read_excel(file)
        all_data = all_data.append(df, ignore_index=True)
    return all_data.drop_duplicates()

def handel_main(filepath_list, title):
    local_time = getDate.getLocalTime()
    df = merge_excel(filepath_list)
    fileabsolutepath = os.path.dirname(
        os.path.realpath(__file__)) + '\\tablefile' + '\\' + title + '\\' + title + local_time + 'total.xlsx'
    df.to_excel(fileabsolutepath, index=False)

    return fileabsolutepath

def send_email_total(title, filepath, receiver, warntitle):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    receivers = receiver
    mail.To = receivers[0]
    mail.Subject = title + '\n'
    mail.Body = title + '\n' + warntitle
    if get_file_size(filepath) < 20:
        try:
            mail.Attachments.Add(filepath)  # 获取本文件路径 os.path.dirname(os.path.realpath(__file__)) + '\\' + filename
        except Exception as e:
            mail.Body = title + u'发送失败！' + '\n' + str(e)  # e.value
            logger.error(e)
        mail.Send()
        print(title + u'邮件发送成功！')
        logger.info(title + u'邮件发送成功！')
    else:
        mail.Body = title + u'文件大小超过20M,发送失败！'
        mail.Send()
        logger.info(title + u'文件大小超过20M,发送失败！')

def get_file_size(filePath):
    fsize = os.path.getsize(filePath)
    fsize = fsize / float(1024 * 1024)
    return round(fsize, 2)



