# -*- coding: utf-8 -*-
"""
Created on Sat Jan 16 14:54:23 2021

@author: wanpeng.xie
"""

import pandas as pd
import numpy as np
import time

# addr = 'C:\\Users\\wanpeng.xie\\Desktop\\主数据处理.xlsx'
# masterData = pd.read_excel(addr, sheet_name=None)
# mainDataZ8 = masterData['主数据']
# handAccount = masterData['手工账']
# zpp043a = masterData['zpp043a']
# changeData = masterData['成品改工厂']
# sapData = masterData['成品主数据']
# settleMainData = pd.DataFrame(np.arange(3).reshape((1,3)),columns=['编码', '物料描述', '工厂处理'])
# changeData.insert(11, '修改', ' ')

class changgeFactory:
    
    def importFiles(self, ):
        addr = 'C:\\Users\\wanpeng.xie\\Desktop\\主数据处理.xlsx'
        global masterData, mainDataZ8, handAccount, zpp043a, changeData, sapData, settleMainData
        masterData = pd.read_excel(addr, sheet_name=None)
        mainDataZ8 = masterData['主数据']
        handAccount = masterData['手工账']
        zpp043a = masterData['zpp043a']
        changeData = masterData['成品改工厂']
        sapData = masterData['成品主数据']
        settleMainData = pd.DataFrame(np.arange(3).reshape((1,3)),columns=['编码', '物料描述', '工厂处理'])
        changeData.insert(11, '修改', ' ')
    
    def checkZ1(self, ):
        global changeData, sapData
        sapData = sapData.applymap(str)
        changeData = changeData.applymap(str)
        print(type(sapData.loc[0, '特定工厂状态']))
        
        changeDict = {}
        for i in range(len(changeData)):
            changeDict[str(changeData.loc[i, '成品编码'])] = str(changeData.loc[i, '改后工厂'])
        sapData = sapData.applymap(str)
        sapData = sapData.loc[sapData['工厂'].str.contains('2022|2023')]
        
        #开始核查主数据
        i = 0
        for code in changeDict:
            tempIndex = list(sapData.loc[sapData['物料编码'].str.contains(code)].index)
            #先看一下是否需要扩充BOM
            factory = list(sapData.loc[tempIndex, '工厂'])
            print(changeDict[code], type(changeDict[code]), factory)
            if changeDict[code] not in factory:
                settleMainData.loc[i, '编码'] = code
                settleMainData.loc[i, '物料描述'] = sapData.loc[tempIndex[0], '物料描述']
                settleMainData.loc[i, '工厂处理'] = '扩充BOM' + ',' + changeDict[code] + '工厂'
                changeDataIndex = changeData.loc[changeData['成品编码'].str.contains(code)].index
                changeData.loc[changeDataIndex, '修改'] = '扩充BOM' + ',' + changeDict[code]
                i = i+1
            else:
               print(code, '两个工厂都有BOM，无需扩充')
            
            #再解决一下删除标识和特定工厂状态
            for index in tempIndex:
                if sapData.loc[index, '删除标识'] == 'X' and sapData.loc[index, '特定工厂状态'] != 'nan' and sapData.loc[index, '工厂'] == changeDict[code]:
                    settleMainData.loc[i, '编码'] = code
                    settleMainData.loc[i, '物料描述'] = sapData.loc[tempIndex[0], '物料描述']
                    settleMainData.loc[i, '工厂处理'] = '取消X' + ',' + sapData.loc[index, '特定工厂状态'] + '标识' + ',' + sapData.loc[index, '工厂'] + '工厂'
                    changeDataIndex = changeData.loc[changeData['成品编码'].str.contains(code)].index
                    changeData.loc[changeDataIndex, '修改'] = '取消X' + ',' + sapData.loc[index, '特定工厂状态'] + '标识' + ',' + sapData.loc[index, '工厂'] + '工厂'
                    i = i+1
                    continue
                elif sapData.loc[index, '删除标识'] == 'X' and sapData.loc[index, '工厂'] == changeDict[code]:
                    settleMainData.loc[i, '编码'] = code
                    settleMainData.loc[i, '物料描述'] = sapData.loc[tempIndex[0], '物料描述']
                    settleMainData.loc[i, '工厂处理'] = '取消删除标识' + ',' + sapData.loc[index, '工厂'] + '工厂'
                    changeDataIndex = changeData.loc[changeData['成品编码'].str.contains(code)].index
                    changeData.loc[changeDataIndex, '修改'] = '取消删除标识' + ',' + sapData.loc[index, '工厂'] + '工厂'
                    i = i+1
                    continue
                elif sapData.loc[index, '特定工厂状态'] != 'nan' and sapData.loc[index, '工厂'] == changeDict[code]:
                    settleMainData.loc[i, '编码'] = code
                    settleMainData.loc[i, '物料描述'] = sapData.loc[tempIndex[0], '物料描述']
                    settleMainData.loc[i, '工厂处理'] = '取消' + sapData.loc[index, '特定工厂状态'] + '标识' + ',' + sapData.loc[index, '工厂'] + '工厂'
                    changeDataIndex = changeData.loc[changeData['成品编码'].str.contains(code)].index
                    changeData.loc[changeDataIndex, '修改'] =  '取消' + sapData.loc[index, '特定工厂状态'] + '标识' + ',' + sapData.loc[index, '工厂'] + '工厂'
                    i = i+1
                    continue
                elif sapData.loc[index, '工厂'] == changeDict[code]:
                    changeDataIndex = changeData.loc[changeData['成品编码'].str.contains(code)].index
                    changeData.loc[changeDataIndex, '修改'] = '主数据没有问题无需修改'
                    continue

    def saveFiles(self, ):
        fileName = '改工厂主数据处理' + time.strftime("%Y-%m-%d", time.localtime()) + '.xlsx'
        addr = 'C:\\Users\\wanpeng.xie\\Desktop\\主数据处理\\' + fileName
        with pd.ExcelWriter(addr) as writer:
            changeData.to_excel(writer, sheet_name='改工厂')
            settleMainData.to_excel(writer, sheet_name='主数据')      

if __name__ == '__main__':
    C = changgeFactory()
    C.importFiles()
    C.checkZ1()
    C.saveFiles()
