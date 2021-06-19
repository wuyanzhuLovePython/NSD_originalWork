# -*- coding: utf-8 -*-
"""
Created on Thu Jun 17 20:50:55 2021

@author: wanpeng.xie
"""

import pandas as pd
import numpy as np



def findIndex(report):
    if report['仓储地点的描述'] == 0:
        print(report.name)
        tempIndex.append(report.name)

def calcSum(report):
    
    return (report['非限制使用的估价的库存'] + report['质量检验中的库存'] + report['冻结的库存'])
        
#汇总表的物料组描述进行改变
def totalChange(report):
    if report['产品中类描述'] == '回收空硒鼓':
        report['物料组描述'] = '空壳'
    elif report['产品中类描述'] == '自制半成品':
        report['物料组描述'] = '自制半成品'
    elif report['产品中类描述'] == '外购半成品':
        report['物料组描述'] = '半成品外购'
    
    return report  

#将透视的格式转为常规的Dataframe格式
def changePivot(pivotTemp):
    lengths = len(pivotTemp)#获取长
    wides = len(pivotTemp.keys()) + 1#获取宽
    tempcolumns = [pivotTemp.index.name] + list(pivotTemp.keys()) #获取列
    print(lengths, wides, tempcolumns)
    tempData = pd.DataFrame(np.arange(lengths*wides).reshape((lengths, wides)), columns=tempcolumns)
    
    flag = 0
    for column in tempcolumns:
        if flag == 0:
            tempData.loc[:, column] = list(pivotTemp.index)
        else:
            tempData.loc[:, column] = list(pivotTemp.loc[:, column])
            
        flag += 1
    
    return tempData

def fillFamily(report):
    if report['大类'] == '自制半成品':
        report['FAMILY_DESCR'] = '自制半成品'
    elif report['大类'] == '空壳':
        report['FAMILY_DESCR'] = '空壳'
    
    return report

def modifyUnused(report):
    
    if report['ITEM_ID'] in assisList:
        report['UNUSEDQTY'] = report['UNUSEDQTY']/10000
        print('要处理的辅助件编码： ',report['ITEM_ID'])
    else:
        report['UNUSEDQTY'] = report['UNUSEDQTY']
    return report

#获取字典
def createDict(report, a, b, dictory):
    dictory[report[a]] = report[b]
    
    return dictory

if __name__ == '__main__':
    
    print('正在读取文件。。。。')
    addr = 'C:\\Users\\wanpeng.xie\\Desktop\\自动非分配报表.xlsx'
    data = pd.read_excel(addr, sheet_name=None)
    print('读取成功！')
    
    #删除辅助件的两行乱码    
    tempIndex = []
    data['注塑件和五辅件'].apply(findIndex, axis=1)
    tempIndex.append(tempIndex[0]-1)
    data['注塑件和五辅件'].drop(index=tempIndex, inplace=True)
    
    #删除重复的外购注塑件整套
    tempIndex = data['大件报表'].loc[data['大件报表']['产品中类描述'].str.contains('外购注塑件整套')].index
    data['大件报表'].drop(index=tempIndex, inplace=True)
    
    #合并三个表，得到总体库存表
    totalData = pd.concat([data['大件报表'], data['全新易耗件'], data['注塑件和五辅件']], ignore_index=True)
    totalData.insert(list(totalData.keys()).index('冻结的库存'), '数量', 0)
    totalData['数量'] = totalData.apply(calcSum, axis=1)
    totalData = totalData.apply(totalChange, axis=1)#汇总表的物料组描述进行改变
    
    #透视库存总表
    tempTotal = totalData.pivot_table(index=['物料组描述'], values = ['数量'], aggfunc={'数量':np.sum})
    tempTotal = changePivot(tempTotal)
    
    #处理未分配表
    data['未分配原始表'] = data['未分配原始表'].loc[data['未分配原始表']['LOC_ID'] != '2023/PO']
    data['未分配原始表'] = data['未分配原始表'].apply(fillFamily, axis=1)
    
    #先把辅助件挑出来，然后除以10000
    data['辅助件']['采购订单单位'] = data['辅助件']['采购订单单位'].apply(str)
    data['辅助件'] = data['辅助件'].loc[data['辅助件']['采购订单单位'].str.contains('M2')]
    assisList = list(data['辅助件']['物料编码'])
    assisList = [str(i) for i in assisList]
    data['未分配原始表']['ITEM_ID'] = data['未分配原始表']['ITEM_ID'].apply(str)
    data['未分配原始表'].apply(modifyUnused, axis=1)
    
    #透视未分配报表
    tempUnsed = data['未分配原始表'].pivot_table(index=['FAMILY_DESCR'], values=['UNUSEDQTY'], aggfunc={'UNUSEDQTY': np.sum}) 
    tempUnsed = changePivot(tempUnsed)

    #将透视的内容转化成字典
    unusedNumDict = {}
    unusedNumDict = tempUnsed.apply(createDict, axis=1, args=(tempUnsed.keys()[0], tempUnsed.keys()[1], unusedNumDict))[0]
    
    #处理汇总的表
    data['汇总'].insert(list(data['汇总'].keys()).index('5-28未分配'), '6-17未分配', 0)
    data['汇总']['6-17未分配'] = data['汇总']['物料描述'].apply(lambda x : unusedNumDict[x] if (x in unusedNumDict) else '不存在')
    
    #将将总库存转化成字典
    totalDict = {}
    totalDict = tempTotal.apply(createDict, axis=1, args=(tempTotal.keys()[0], tempTotal.keys()[1], totalDict))[0]
    
    #总库存放入汇总
    data['汇总']['总库存'] = data['汇总']['物料描述'].apply(lambda x : totalDict[x] if (x in totalDict) else '不存在')
    
    
    
    
