# -*- coding: utf-8 -*-
"""
Created on Thu May 27 22:07:38 2021

@author: 吴彦祖的laptop
"""
# -*- coding: utf-8 -*-
"""
Created on Wed Oct 28 08:43:14 2020

@author: wanpeng.xie
"""

import pandas as pd
import numpy as np
import openpyxl as op
import time



class deliveryAndTax:
    
    def importFile(self, ):
        addr = 'C:\\Users\\wanpeng.xie\\Desktop\\TAX.xlsx'
        global planData, sapData, delivery, payTax, newPayTax, newPlanData
        
        planData = pd.read_excel(addr, 'Sheet1')
        sapData = pd.read_excel(addr, 'Sheet2')
        delivery = pd.DataFrame(np.arange(13).reshape((1,13)), columns=['序号','转出物料号', '转出物料描述', '转出单位', '转出数量',  '转出仓位', '转入物料号', '转入物料描述', '转入单位', '转入数量', '转入仓位', '转入工厂', '批次'])
        payTax = pd.read_excel(addr, 'Sheet4')
        newPayTax = pd.DataFrame(np.arange(len(list(payTax.columns))).reshape((1,len(list(payTax.columns)))), columns=list(payTax.columns))
        newPlanData = pd.DataFrame(np.arange(5).reshape((1,5)), columns=['半成品编码','物料描述', '数量', '仓位', '批次'])

    
    
    def mergePlan(self,):
        # addr = 'E:/autoWork/TAX.xlsx'
        
        global planData, sapData, delivery, payTax, newPayTax, newPlanData
        # planData = pd.read_excel(addr, 'Sheet1')
        # sapData = pd.read_excel(addr, 'Sheet2')
        # delivery = pd.DataFrame(np.arange(13).reshape((1,13)), columns=['序号','转出物料号', '转出物料描述', '转出单位', '转出数量',  '转出仓位', '转入物料号', '转入物料描述', '转入单位', '转入数量', '转入仓位', '转入工厂', '批次'])
        # payTax = pd.read_excel(addr, 'Sheet4')
        # newPayTax = pd.DataFrame(np.arange(len(list(payTax.columns))).reshape((1,len(list(payTax.columns)))), columns=list(payTax.columns))
        # newPlanData = pd.DataFrame(np.arange(5).reshape((1,5)), columns=['半成品编码','物料描述', '数量', '仓位', '批次'])
        
        planData['半成品编码'] = planData['半成品编码'].apply(str)
        planData['批次'] = planData['批次'].apply(str)
        planData['仓位'] = planData['仓位'].apply(str)
        codeList = list(set(planData.loc[:,'半成品编码']))
        print('成品编码：', codeList)
        
        i = 0
        string = ''
        for code in codeList:
            tempList = list(planData.loc[planData['半成品编码'].str.contains(code)].index)
            print('index:',tempList)
            num = planData.loc[tempList, '数量'].sum()
            print('数量：',num)
            PositionLength = len(set(list(planData.loc[tempList, '仓位'])))
            Positions = list(set(list(planData.loc[tempList, '仓位'])))
            print('当前仓位重复数:', PositionLength)
            print(Positions)
            
            #把不同仓位的区分出来
            
            if PositionLength > 1:
                for positon in Positions:
                    PositionData = planData.loc[tempList, :]
                    PositionIndex = list(PositionData.loc[PositionData['仓位'].str.contains(positon)].index)
                    print(PositionIndex)
                    newPlanData.loc[i, '半成品编码'] = code
                    newPlanData.loc[i, '物料描述'] = PositionData.loc[PositionIndex[0], '物料描述']
                    num = planData.loc[PositionIndex, '数量'].sum()
                    newPlanData.loc[i, '数量'] = num
                    newPlanData.loc[i, '仓位'] = PositionData.loc[PositionIndex[0], '仓位']
                    
                    listLength = len(list(PositionData.loc[PositionIndex, '批次']))
                    
                    if listLength > 1:
                        string = ' '.join(list(PositionData.loc[PositionIndex, '批次']))
                    else:
                        string = list(PositionData.loc[PositionIndex, '批次'])
                    
                    newPlanData.loc[i, '批次'] = string
                    i = i+1
                continue
                            
            
            newPlanData.loc[i, '半成品编码'] = code
            newPlanData.loc[i, '物料描述'] = planData.loc[tempList[0], '物料描述']
            newPlanData.loc[i, '数量'] = num
            newPlanData.loc[i, '仓位'] = planData.loc[tempList[0], '仓位']
            
            listLength = len(list(planData.loc[tempList, '批次']))
            print('当前半成品编码重复数:', listLength)
            
            if listLength > 1:
                string = ' '.join(list(planData.loc[tempList, '批次']))
            else:
                string = list(planData.loc[tempList, '批次'])
            
            newPlanData.loc[i, '批次'] = string
            i = i+1
        planData = newPlanData
        print('--------mergePlan函数运行完毕----------', '\n')
    
    def PlanToDe(self,):
        length = len(planData)
        for i in range(length):
            delivery.loc[i, '转出物料号'] = planData.loc[i, '半成品编码']
            delivery.loc[i, '转入物料号'] = planData.loc[i, '半成品编码']
            delivery.loc[i, '转出物料描述'] = planData.loc[i, '物料描述']
            delivery.loc[i, '转入物料描述'] = planData.loc[i, '物料描述']
            delivery.loc[i, '转出数量'] = planData.loc[i, '数量']
            delivery.loc[i, '转入数量'] = planData.loc[i, '数量']
            delivery.loc[i, '转出单位'] = 'PCS'
            delivery.loc[i, '转入单位'] = 'PCS'
            delivery.loc[i, '转出仓位'] = planData.loc[i, '仓位']

            if int(planData.loc[i, '仓位']) < 2000:
                delivery.loc[i, '转入仓位'] = int(planData.loc[i, '仓位']) + 1000
            else:
                delivery.loc[i, '转入仓位'] = int(planData.loc[i, '仓位'])
            delivery.loc[i,['批次']] = planData.loc[i, '批次']
            delivery.loc[i, '转入工厂'] = '2023'
        
     
     
    def sapToTax(self,):
        global sapData, payTax, planData
        sapData['父件编码'] = sapData['父件编码'].apply(str)
        
        #把核心的半成品编码提取出来
        planLength = len(planData)
        sapLenght = len(sapData)
        halfGoodCode = []
        for index in range(planLength):
            halfGoodCode.append(planData.loc[index, '半成品编码']) 
        print('sapToTax函数的半成品编码:',halfGoodCode)
        
        #以半成品编码为核心，对sapData进行遍历和筛选
        i = 0
        #从原始表一行一行过，看每一行需要补多少税，最后再合并
        for pIndex in range(len(planData)):
            for index in range(sapLenght):
                if sapData.loc[index, '父件编码'] == planData.loc[pIndex, '半成品编码']:
                    subCode = sapData.loc[index, '子件编码']
                    subDesc = sapData.loc[index, '子件描述']
        
                    payTax.loc[i, '物料号'] = str(subCode)
                    payTax.loc[i, '物料描述'] = subDesc
                    #print(planData.loc[i, '数量'])
                    
                    #专门解决数量这一列
                    #处理主要需要消耗的碳粉
                    if sapData.loc[index, '子件单位'] == 'KG':
                        num = sapData.loc[index, '子件用量']
                        payTax.loc[i, '数量'] = num * float(planData.loc[pIndex, '数量'])
                        payTax.loc[i, '单位'] = 'KG'
                        print('KG:' ,num, planData.loc[pIndex, '数量'])
                    
                    #处理按g的碳粉
                    elif sapData.loc[index, '子件单位'] == 'G':
                        num = float(sapData.loc[index, '子件用量']) * 0.001
                        payTax.loc[i, '数量'] = num * float(planData.loc[pIndex, '数量'])
                        payTax.loc[i, '单位'] = 'KG'
                        print('G:',num, planData.loc[pIndex, '数量'])
                        
                        
                    #其他的基本上是和数量是一一对应的
                    else:
                        num = sapData.loc[index, '子件用量']
                        payTax.loc[i, '数量'] = num * planData.loc[pIndex, '数量']
                        payTax.loc[i, '单位'] = 'PCS'
                    i = i + 1
    
    def mergeDupl(self,):
        #prepared action
        global payTax
        payTax['物料号'] = payTax['物料号'].apply(str)
        codeList = list(set(payTax['物料号']))
        print(codeList)
        i = 0
        
        #开始求和
        for code in codeList:
            tempList = list(payTax.loc[payTax['物料号'].str.contains(code)].index)#把所有的相同物料编码的都找出来
            num = payTax.loc[tempList, '数量'].sum() #把这些物料按照index集合起来求和
            print(num)
            newPayTax.loc[i, '数量'] = num
            newPayTax.loc[i, '物料号'] = code
            newPayTax.loc[i, '物料描述'] = payTax.loc[tempList[0], '物料描述']#反正templist中装的都是同样的物料号和编码
            newPayTax.loc[i, '单位'] = payTax.loc[tempList[0], '单位']
            i = i + 1
              
    def saveTax(self,):
        
        addr = 'C:\\Users\\wanpeng.xie\\Desktop\\补税模板.xlsx'
        
        #操作转储部分
        excel = op.load_workbook(addr)
        s = excel['转储']['B7':'M52']
        rowsLen = len(s[0])
        colLen = len(s)
        print(rowsLen, colLen)
        
        print('--------------测试OPenPYXL-----------------------')
        print(len(delivery.loc[:, '转出物料号':'批次']))
        headList = list(delivery.loc[:, '转出物料号':'批次'].keys())
        testList = []
        
        delivery.loc[:, '转出物料号':'批次'].apply(lambda x : [testList.append(x[col]) for col in headList], axis=1)
        #删除重复的第一行
        testList = testList[len(delivery.loc[:, '转出物料号':'批次'].keys()):]
        
        print('转储列表内容：',testList)
        
        #设置一下时间
        excel['转储']['J5'].value = time.strftime('%Y-%m-%d', time.localtime())
        #开始赋值,只对转储
        flag = 0
        for i in s:
            for j in i:
                j.value = testList[flag]
                print(j, j.value)
                if testList[flag] == testList[-1]:
                    break
                flag += 1
            #跳出循环后会执行下面的if语句，把跳出循环的条件再次执行就OK 
            if testList[flag] == testList[-1]:
                break
        
        #操作补税部分
        v = excel['补税']['B6':'E27']
        taxHeadList = list(newPayTax.loc[:, '物料号':'数量'])
        testList = []
        
        #这个是超级代码，函数式编程终极使用。但美中不足是会重复加第一行。。。。。而且目前还没找到问题
        newPayTax.loc[:, '物料号':'数量'].apply(lambda x : [testList.append(x[col]) for col in taxHeadList], axis=1)
        #删除重复的第一行
        testList = testList[len(newPayTax.loc[:, '物料号':'数量'].keys()):]#左闭右开，填-1取不到最后一个。
        print('补税列表内容：',testList)
        
        #设置时间
        excel['补税']['H4'].value = time.strftime('%Y-%m-%d', time.localtime())
        
        flag = 0
        for i in v:
            for j in i:
                j.value = testList[flag]
                print(j, j.value)
                if testList[flag] == testList[-1]:
                    break
                flag += 1
            #跳出循环后会执行下面的if语句，把跳出循环的条件再次执行就OK 
            if testList[flag] == testList[-1]:
                break
        
        halfGood = ''
        toner = ''
        drum = ''
        
        for cell in list(delivery.loc[:, '转出物料描述']):
            if '半成品' in cell:
                halfGood = '半成品'
            elif '碳粉' in cell:
                toner = '碳粉'
            elif '感光鼓' in cell:
                drum = '感光鼓'
        
        fileName = time.strftime('%Y-%m-%d', time.localtime()) + '' + halfGood + toner + drum + '.xlsx'
        
        saveAddr = 'C:\\Users\\wanpeng.xie\\Desktop\\补税\\2021年七月\\' + fileName
        excel.save(saveAddr)


if __name__ == '__main__':
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
    #input()
    
