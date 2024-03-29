# -*- coding: utf-8 -*-
"""
Created on Mon May 31 22:45:02 2021

@author: 吴彦祖的laptop
"""

import pandas as pd
import openpyxl as vb



def importFiles():
    
    addr = 'C:\\Users\\吴彦祖的laptop\\Desktop\\company\\2021年（6）月第（4周） 运营资料-2021.6.22-总.xlsx'
    
    
    data = pd.read_excel(addr, sheet_name=None)
    print(data.keys())
    
    totalReport = data['总表']
    client = data['客户']
    regionalSeries = data['区域系列']
    clientSeries = data['系列客户']
    keyProducts = data['2021年重点产品']
    
    return (totalReport, client, regionalSeries, clientSeries, keyProducts)

def importExcel():
    
    addr = 'C:\\Users\\吴彦祖的laptop\\Desktop\\company\\2021年（6）月第（4周） 运营资料-2021.6.22-总.xlsx'
    
    vbExcel = vb.load_workbook(addr)
    
    return vbExcel

def creatDict(deleteList, dictDele):
    
    tempList = []
    tempList.append(deleteList[0])
    for i in deleteList:
        try:
            nextDel = deleteList[deleteList.index(i) + 1]#nextDel 是列表当前遍历元素的下一个
        except:
            tempList.append(i)
        #print(i, nextDel)
    
        if i - nextDel > 1:
            tempList.append(i)
            tempList.append(nextDel)
            
        #将列表转化成字典    
        for i in tempList:
            if i == tempList[-1]:
                break
            if len(tempList[0:tempList.index(i)])%2 == 0:
                dictDele[i] = tempList[tempList.index(i)+1]     
                
    return dictDele    

def dealTotalReport(operationExcel, index):
    
    totalSheet = operationExcel['总表']
    #处理总表
    
    # if rigionName == '德国区':
    #     continue
    #获取excel表中的row数据
    #获取准确的row数据
    colunms = totalSheet['B']
    headTempList = []
    lastTempList = []
    for i in colunms:
        if i.value == '单品毛利$':
            lastTempList.append(i.row)
        elif i.value == '数量':
            headTempList.append(i.row)
    
    contentGap = 3 + (lastTempList[0] - headTempList[0] + 1) #3是表头长度，列表相减是内容长度
    
    headRow = index
    lastRow = index + contentGap
    print(headRow, lastRow, contentGap)
    
    totalSheet.delete_rows(lastRow+1, lastTempList[-1])
    if headRow == 2:
        pass
    else:
        totalSheet.delete_rows(headTempList[0], headRow-1)
        
def clientReport(operationExcel, index, rigionName):
    
    #处理客户表
    #规律 excel的索引 == index + 2
    clientSheet = operationExcel['客户']
    client[1] = client[1].apply(str)
    client[0] = client[0].apply(str)
    tempindex = client.loc[client[1].str.contains(rigionName)].index
    if len(tempindex) == 0:
        tempindex = client.loc[client[0].str.contains(rigionName)].index
    print(tempindex, rigionName, '客户')
    #获取excel表中的row数据
    clientList = list(range(5, len(client)+2))
    #获取excel表中的row数据
    if len(tempindex):
        
        print('以下是删除的列表')
        deleteList = list(tempindex)#需要删除的index
        print(deleteList)
        deleteList = [i+2 for i in deleteList]#将dataframe的格式转化为excel的样子
        print(deleteList)
        deleteList = list(set(clientList) - set(deleteList))#剔除不该删除的
        deleteList.sort(reverse=True)#设置顺序，从后往前删除

        
        print('删除的为以上，客户')

        
        delDict = {}
        delDict = creatDict(deleteList, delDict)
        print('当前删除的列表为：',delDict)
        
        for head, last in delDict.items():
            if last == len(client)+2 :
                clientSheet.delete_rows(last, head)
            else:
                head = head - (last-1)
                clientSheet.delete_rows(last, head)
            
            
    else:
        pass

    print(rigionName, '客户', '运行成功！')
    
def clientSeriesReport(operationExcel, index, rigionName):
    
    #处理系列客户
    cliSeriesSheet = operationExcel['系列客户']
    clientSeries[0] = clientSeries[0].apply(str)
    clientSeries[1] = clientSeries[1].apply(str)
    tempindex = clientSeries.loc[clientSeries[1].str.contains(rigionName)].index
    if len(tempindex) == 0:
        tempindex = client.loc[client[0].str.contains(rigionName)].index
    #获取excel表中的row数据
    clientSeriesList = list(range(5, len(clientSeries)+2))
    #获取excel表中的row数据
    if len(tempindex):
        
        print('以下是删除的列表')
        deleteList = list(tempindex)#需要删除的index
        print(deleteList)
        deleteList = [i+2 for i in deleteList]#将dataframe的格式转化为excel的样子
        print(deleteList)
        deleteList = list(set(clientSeriesList) - set(deleteList))#剔除不该删除的
        deleteList.sort(reverse=True)#设置顺序，从后往前删除

        
        
        
        delDict = {}
        delDict = creatDict(deleteList, delDict)
        print('当前删除的列表为：',delDict)
        
        for head, last in delDict.items():
            if last == len(clientSeries)+2 :
                cliSeriesSheet.delete_rows(last, head)
            else:
                head = head - (last-1)
                cliSeriesSheet.delete_rows(last, head)

    else:
        pass


    print(rigionName, '系列客户', '运行成功！')
    
def keyProductsReport(operationExcel, rigionName):
    
    keyProdSheet = operationExcel['2021年重点产品']
    keyProducts[1] = keyProducts[1].apply(str)
    keyProducts[0] = keyProducts[0].apply(str)
    tempindex = keyProducts.loc[keyProducts[0].str.contains(rigionName)].index
    if len(tempindex) == 0:
        tempindex = keyProducts.loc[keyProducts[1].str.contains(rigionName)].index
    print(tempindex)
    
    keyProductsList = list(range(4, len(keyProducts)+2))
    #获取excel表中的row数据
    if len(tempindex):
        
        print('以下是删除的列表')
        deleteList = list(tempindex)#需要删除的index
        print(deleteList)
        deleteList = [i+2 for i in deleteList]#将dataframe的格式转化为excel的样子
        print(deleteList)
        deleteList = list(set(keyProductsList) - set(deleteList))#剔除不该删除的
        deleteList.sort(reverse=True)#设置顺序，从后往前删除

        
        
        #重点产品不需要区域删除
       
        for i in deleteList:
            keyProdSheet.delete_rows(i)
    else:
        pass
    
def regionalSeriesReport(operationExcel, rigionName):
    
    regionalSerSheet = operationExcel['区域系列']
    regionalSeries[1] = regionalSeries[1].apply(str)
    regionalSeries[0] = regionalSeries[0].apply(str)
    tempindex = regionalSeries.loc[regionalSeries[1].str.contains(rigionName)].index
    if len(tempindex) == 0:
        tempindex = regionalSeries.loc[regionalSeries[0].str.contains(rigionName)].index
    print(tempindex, len(tempindex), rigionName)
    
    #获取excel表中的row数据
    regionalSeriesList = list(range(5, len(regionalSeries)+2))
    #获取excel表中的row数据
    if len(tempindex):
        
        print('以下是删除的列表')
        deleteList = list(tempindex)#需要删除的index
        print(deleteList)
        deleteList = [i+2 for i in deleteList]#将dataframe的格式转化为excel的样子
        print(deleteList)
        deleteList = list(set(regionalSeriesList) - set(deleteList))#剔除不该删除的
        deleteList.sort(reverse=True)#设置顺序，从后往前删除
        
        
        
        delDict = {}
        delDict = creatDict(deleteList, delDict)
        print('当前删除的列表为：',delDict)
        
        for head, last in delDict.items():
            head = head - (last-1)
            regionalSerSheet.delete_rows(last, head)
                    
            
    else:
        pass

def spliteReports():
    
    
    #修改表头
    reportsList = [totalReport, client, regionalSeries, clientSeries, keyProducts]
    for i in reportsList:
        i.columns = list(range(len(i.keys())))
    
    #分解总表
    totalReport[1] = totalReport[1].apply(str)
    tempindex = totalReport.loc[totalReport[1].str.contains('类别')].index
    print(list(tempindex))
    indexList = list(tempindex)
    
    seriesList = ['亚太区', '英国区', '俄罗斯区', '日本区', '东亚', '中东非', '南美区', 'SCC', '澳洲', '俄意英', 'SCC-ABR', '东欧区', '德国区', '意西区', '北美区', '德东法', '法荷区'] 
    print(seriesList, len(seriesList))
    
    reportsList = []
    for index in indexList:
        
        operationExcel = importExcel()
        rigionName = totalReport.loc[index+3, 2]#获取当前区域
        
        #处理总表
        dealTotalReport(operationExcel, index)
        
        #处理客户表
        clientReport(operationExcel, index, rigionName)


        
        #处理系列客户
        clientSeriesReport(operationExcel, index, rigionName)
        

        
        #重点产品
        keyProductsReport(operationExcel, rigionName)
        
        

        
        #区域系列
        regionalSeriesReport(operationExcel, rigionName)
        
        
        addr = 'D:\\运行结果\\' + rigionName + '.xlsx'
        operationExcel.save(addr)
        
        
        
        
 
        

# def saveFiles1(report, sheetName):
    
#     fileName = '运营资料' + time.strftime("%Y-%m-%d", time.localtime()) + '.xlsx'
#     addr = 'D:\\运行结果\\' + fileName
#     j = 0
#     with pd.ExcelWriter(addr) as writer:
#         for i in report:
#             i.to_excel(writer, sheet_name=str(sheetName[j]))
#             j += 1

    
    
if __name__ == '__main__':
    (totalReport, client, regionalSeries, clientSeries, keyProducts) = importFiles()
    
    spliteReports()
