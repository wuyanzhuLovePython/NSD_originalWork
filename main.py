# -*- coding: utf-8 -*-
"""
Created on Fri Feb 19 20:24:46 2021

@author: wanpeng.xie
"""

#import autoOrder
import os
# import pandas as pd
# import numpy as np
mainAddr = 'C:\\Users\\wanpeng.xie\\Desktop\\自动化'
os.chdir(mainAddr)

from autoPayTax import deliveryAndTax
from autoOrder import autoOrder
from mainData import changgeFactory
from mainUI import maintainUI
from checkShortage import checkShortage
from autoPoDate import maintainPodate
from prettytable import PrettyTable

# mainAddr = 'C:\\Users\\wanpeng.xie\\Desktop\\自动化'
# os.chdir(mainAddr)
# print(os.getcwd())



i=10
for i in range(i):
    
    i=i+1
    
    for i in range(8):
        print("     " * (7-i) + "*  " * i + "  " * i * 2 + "\b" * 7 + "*  " * i)
    for i in range(6,0,-1):
        print("     " * (7-i) + "*  " * i + "  " * i * 2 + "\b" * 7 + "*  " * i)
    
    print("==============欢迎使用自动化工作平台======================")
    
    x = PrettyTable()
    x.field_names = ["工作内容", "KEY"]
    x.add_row(['自动补税', 1])
    x.add_row(['自动改工厂', 2])
    x.add_row(['自动抛单', 3])
    x.add_row(['自动维护UI', 4])
    x.add_row(['自动核对欠料表', 5])
    x.add_row(['自动维护PO交期', 6])
    x.align['工作内容'] = 'l'
    print(x)
    # print(" 补税请输入‘1’，\n 改工厂请输入‘2’, \n 自动抛单请输入‘3’,\
    #       \n 自动维护UI请输入‘4’, \n 自动核对欠料表请输入‘5’, \n 自动维护PO交期请输入‘6’")
    a = int(input("请输入你的需求："))
    
    if a == 1:
       
        print('------------请注意-------------------')
        print('----输入的成品编码不能有重复！------')
        D = deliveryAndTax()
        D.importFile()
        D.mergePlan()
        D.PlanToDe()
        D.sapToTax()
        D.mergeDupl()
        D.saveTax()
        print('更改完成！')
    
    elif a == 2:
        C = changgeFactory()
        C.importFiles()
        C.checkZ1()
        C.saveFiles()
        
    elif a == 3:
        print('------------正导入数据---------------')
        A = autoOrder()
        A.importFiles()
        print('----------正处理数据-------------')
        A.settleSheet()
        print('---------保存文件----------')
        A.savefile()
        print('-----------保存成功！----------------')
        
    elif a == 4:
        M = maintainUI()
        print('------------正在导入数据，请等待-------------------')
        M.importFiles()
        print('------------数据导入完成，开始填充表格-------------------')
        M.extract()
        print('------------开始保存-------------------')
        M.saveFiles()
        print('------------保存成功！-------------------')

    
    elif a == 5:
        C = checkShortage()
        print('-------------------正在核对欠料表，请等待--------------------')
        C.importFiles()
        C.files()
        C.shortageOrder()
        C.saveFiles()
        print('-------------------已经核对完成------------------------')
    
    elif a == 6:
        M = maintainPodate()
        M.inputData()
        M.checkData()
        M.saveFiles()
        



