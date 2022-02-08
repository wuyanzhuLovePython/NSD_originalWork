# -*- coding: utf-8 -*-
"""
Created on Tue Mar 23 07:39:43 2021

@author: 吴彦祖的核能laptop！
"""

import pandas as pd
import re

# addr = 'C:\\Users\\86151\\Desktop\\company\\assembly\\艾派克回复交期.xlsx'
# addr1 = 'C:\\Users\\86151\\Desktop\\company\\assembly\\ZMM154批量改交期模板.xls'
# addr2 = 'C:\\Users\\86151\\Desktop\\company\\assembly\\export.xlsx'

# replyDate = pd.read_excel(addr)
# template = pd.read_excel(addr1)
# originalData = pd.read_excel(addr2)


class maintainPodate:
    
    def inputData(self, ):
        addr = 'C:\\Users\\wanpeng.xie\\Desktop\\艾派克回复交期.xlsx'
        # addr1 = 'C:\\Users\\86151\\Desktop\\company\\assembly\\ZMM154批量改交期模板.xls'
        # addr2 = 'C:\\Users\\86151\\Desktop\\company\\assembly\\export.xlsx'
        
        global replyDate, template, originalData, material_shortage, delay_material
        
        replyDate = pd.read_excel(addr, sheet_name='原始')
        template = pd.read_excel(addr, sheet_name='交期模板')
        originalData = pd.read_excel(addr, sheet_name='订单')
        material_shortage = pd.read_excel(addr, sheet_name='欠料表')
        delay_material = pd.read_excel(addr, sheet_name='逾期')

    
    def checkData(self, ):
        
        global originalData
        
        #delete data which include string with '一粒彩'
        tempIndex = originalData.loc[originalData['供应商简称'].str.contains('一粒彩', na=False)].index
        originalData.drop(index=tempIndex, inplace=True)
        
        
        keyList = list(originalData.keys())
        #print(keyList)
        originalData.insert(keyList.index('料号'), '核对', '')
        originalData.insert(keyList.index('数量'),  '艾派克回复交期', '')
        originalData.insert(keyList.index('交货日期'), '差异', '')
        originalData.index = range(len(originalData))
        
        #should format type of int into 'str' before checking
        originalData.drop(index=len(originalData)-1, inplace=True)
        originalData['采购单号'] = originalData['采购单号'].apply(int)
        originalData['料号'] = originalData['料号'].apply(int)
        originalData['数量'] = originalData['数量'].apply(int)
        replyDate['核对1'] = replyDate['核对1'].apply(str)
        
        print(len(originalData))
        
        for i in range(len(originalData)):
            originalData.loc[i, '核对'] = str(originalData.loc[i, '采购单号']) + str(originalData.loc[i, '料号']) + str(originalData.loc[i, '数量'])
            checkNumber = originalData.loc[i, '核对']

            if len(replyDate.loc[replyDate['核对1'].str.contains(checkNumber)]):
                
                tempIndex = replyDate.loc[replyDate['核对1'].str.contains(checkNumber)].index
                originalData.loc[i, '艾派克回复交期'] =replyDate.loc[tempIndex[0], '计划行日期']
                
                diff = pd.to_datetime(originalData.loc[i, '艾派克回复交期']) - pd.to_datetime(originalData.loc[i, '交货日期'])
                diff = int(re.search(r'(.*) days (.*)', str(diff)).group(1))
                originalData.loc[i, '差异'] = diff
            else:
                originalData.loc[i, '差异'] = '未回复交期'
 
        
        
        #未回复交期的也应该加上去
        global originalDataComple
        
        print('赋值前', len(originalData))
        originalDataComple = originalData
        print('赋值后', len(originalData), len(originalDataComple))
        
        #填入模板,要把不必要的都删除
        originalData['差异'] = originalData['差异'].apply(str)
        tempIndex = originalData.loc[originalData['差异'].str.contains('未回复交期', na=False)].index
        
        if len(tempIndex):
            originalData = originalData.drop(index=tempIndex)
            
        originalData1 = originalData[originalData['差异'] != '0']
        originalData1.index = range(len(originalData1))
        
        print('赋值最后', len(originalData), len(originalDataComple), len(originalData1))
        # template.drop(index=len(template), inplace=True)
        print(len(originalData1))
        for i in range(len(originalData1)):
            template.loc[i, '采购单号'] = originalData1.loc[i, '采购单号']
            template.loc[i, '项次'] = originalData1.loc[i, '项次']
            template.loc[i, '计划行号'] = originalData1.loc[i, '计划行号']
            template.loc[i, '交货日期'] = originalData1.loc[i, '艾派克回复交期']
            template.loc[i, '数量'] = originalData1.loc[i, '数量']
            template.loc[i, '操作'] = 'U'
            
    def merge(self, report):
        
        return str(int(report['PO单号'])) + str(int(report['对应物料组件编码'])) + str(int(report['未清数量']))
    
    def merge1(self, report):
        
        return str(int(report['采购单号'])) + str(int(report['料号'])) + str(int(report['数量']))
    
    
    
    def check_effect(self, ):
        
        #先将欠料表处理好
        global material_shortage
        material_shortage = material_shortage.loc[~(material_shortage['PO单号'].isnull())]
        material_shortage.insert(13, 'check', 0)
        material_shortage['check'] = material_shortage.apply(self.merge, axis=1, args=())
        material_shortage['Index'] = list(material_shortage.index)
        
        
        
        #再处理一下逾期的订单
        delay_material.insert(3, 'check', 0)
        delay_material['check'] = delay_material.apply(self.merge1, axis=1, args=())
        delay_list = list(set(delay_material['check']))
        
        global delay_shortage
        delay_shortage = material_shortage
        index_list = []
        delay_shortage.apply(lambda x: index_list.append(x['Index']) if str(x['check']) in delay_list  else x, axis=1)
        
        delay_shortage = delay_shortage.loc[index_list, :]
        
        
        #处理回复交期过慢的订单
        global temp_huge_diff, temp_unreply
        
        temp_unreply = originalDataComple.loc[originalDataComple['差异'] == '未回复交期']
        
        temp_huge_diff = originalDataComple.loc[originalDataComple['差异'] != '未回复交期']
        temp_huge_diff['差异'] = temp_huge_diff['差异'].apply(int)
        temp_huge_diff = temp_huge_diff.loc[temp_huge_diff['差异'] > 2]
        
        #将两个异常合并
        total_abnormal = temp_unreply.append(temp_huge_diff)
        obnormal_list = list(set(total_abnormal.loc[:, '核对']))#挑出核对的索引
        
        #找出欠料表的问题
        index_list = []
        material_shortage.apply(lambda x: index_list.append(x['Index']) if str(x['check']) in obnormal_list  else x, axis=1)
        print(index_list)
        
        material_shortage = material_shortage.loc[index_list, :]
        
        
        
        

            
    def saveFiles(self, ):
        #fileName = '改工厂主数据处理' + time.strftime("%Y-%m-%d", time.localtime()) + '.xlsx'
        global originalDataComple, delay_shortage
        
        addr = 'C:\\Users\\wanpeng.xie\\Desktop\\ZMM154修改.xls' 
        with pd.ExcelWriter(addr) as writer:
            template.to_excel(writer, sheet_name='ZMM154')
            originalDataComple.to_excel(writer, sheet_name='原始数据')   
            material_shortage.to_excel(writer, sheet_name='交期影响的欠料表')
            delay_shortage.to_excel(writer, sheet_name='逾期的欠料表')
        
if __name__ == '__main__':
    
    M = maintainPodate()
    M.inputData()
    M.checkData()
    # M.check_effect()
    M.saveFiles()
