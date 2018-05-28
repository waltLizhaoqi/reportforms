'''
Created on Mar 20, 2018

@author: lizhaoq1
'''
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk
import SalesOfDailyReport
import TheLatestInventory
import PromotionDetails
import BusinessDocument
import PaymentOrder
import ExpensesDocuments

root = tk.Tk()
root.minsize(400, 250)
root.resizable(0, 0)
#root.minsize(400, 250) #400, 250
root.title("Export Report")    # 添加标题

ttk.Label(root, text="Chooes Report", font=("Helvetica", 15, "bold"), borderwidth=3).grid(column=0, row=0, pady=10)    # 添加一个标签，并将其列设置为1，行设置为0


# 复选框
chvarSDR = tk.IntVar()   # 用来获取复选框是否被勾选，通过chVarDis.get()来获取其的状态,其状态值为int类型 勾选为1  未勾选为0
check1 = tk.Checkbutton(root, text="销售日报",variable=chvarSDR, font=('Helvetica', 12))    # text为该复选框后面显示的名称, variable将该复选框的状态赋值给一个变量，当state='disabled'时，该复选框为灰色，不能点的状态
check1.deselect()     # 该复选框是否勾选,select为勾选, deselect为不勾选
check1.grid(column=100, row=1, sticky=tk.W)       # sticky=tk.W  当该列中其他行或该行中的其他列的某一个功能拉长这列的宽度或高度时，设定该值可以保证本行保持左对齐，N：北/上对齐  S：南/下对齐  W：西/左对齐  E：东/右对齐

chvarTLI = tk.IntVar()
check2 = tk.Checkbutton(root, text="最新库存", variable=chvarTLI, font=('Helvetica', 12))
check2.deselect()
check2.grid(column=100, row=2, sticky=tk.W)

chvarPD = tk.IntVar()
check3 = tk.Checkbutton(root, text="促销明细", variable=chvarPD, font=('Helvetica', 12))
check3.deselect()
check3.grid(column=100, row=3, sticky=tk.W)

chvarPO = tk.IntVar()
check3 = tk.Checkbutton(root, text="付款单", variable=chvarPO, font=('Helvetica', 12))
check3.deselect()
check3.grid(column=100, row=4, sticky=tk.W)

chvarED = tk.IntVar()
check3 = tk.Checkbutton(root, text="费用单据", variable=chvarED, font=('Helvetica', 12))
check3.deselect()
check3.grid(column=100, row=5, sticky=tk.W)

chvarBD = tk.IntVar()
check3 = tk.Checkbutton(root, text="业务单据", variable=chvarBD, font=('Helvetica', 12))
check3.deselect()
check3.grid(column=100, row=6, sticky=tk.W)

receiver = ['zhaoqi.li@mattel.com;Kama.Wang@Mattel.com;lu.jiang@mattel.com'] #;Kama.Wang@Mattel.com;lu.jiang@mattel.com
user_list = [('161006031101','123456','16','1006031101'), ('18V08747','123456','18','V08747'),
                ('1558098','123456','15','58098'), ('131006031102','123456','13','1006031102'),
                ('121006031103','123456','12','1006031103'), ('881006031105','123456','88','1006031105'),
                ('321006031107','123456','32','1006031107'), ('311006031110','123456','31','1006031110')]
user_list2 = [('161006031101','123456','16'),('18V08747','123456','18'),
                ('1558098','123456','15'),('131006031102','123456','13'),
                ('121006031103','123456','12'),('881006031105','123456','88'),
                ('321006031107','123456','32'),('311006031110','123456','31')]
def clickMe():
    if chvarSDR.get() == 1:
        SalesOfDailyReport.salesOfDailyReport(user_list2, receiver)
    if chvarTLI.get() == 1:
        TheLatestInventory.theLastInventory(user_list, receiver)
    if chvarPD.get() == 1:
        PromotionDetails.promotionDetails(user_list, receiver)
    if chvarPO.get() == 1:
        PaymentOrder.paymentOrder(user_list, receiver)
    if chvarED.get() == 1:
        ExpensesDocuments.expensesDocuments(user_list, receiver)
    if chvarBD.get() == 1:
        BusinessDocument.businessDocument(user_list, receiver)

action = ttk.Button(root, text="Export", command=clickMe)
action.grid(column=0, row=7, ipadx=20, pady=10)
ttk.Button(root, text='Quit', command=root.quit).grid(column=250, row=7, ipadx=20, pady=10)

root.mainloop()      # 当调用mainloop()时,窗口才会显示出来


