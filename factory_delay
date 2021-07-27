# -*- coding: utf-8 -*-
"""
Created on Mon Jul 26 22:59:05 2021

@author: 吴彦祖的laptop
"""

import pandas as pd


class delayReport:
    
    def __init__(self, ):
        self.addr = r'C:\Users\吴彦祖的laptop\Desktop\company\自动欠料表核对.xlsx'
        self.data = pd.read_excel(self.addr, sheet_name=None)
        self.delay_report = self.data['工厂延期']
        
    def operate_delay(self, ):
        
        original_data = self.data['Sheet1'] 
        print(original_data)
        
        
        original_data['协助人'] = original_data['协助人'].apply(str)#只有变成string才能赛选，而且这个还能对付空值
        original_data = original_data.loc[original_data['协助人'].str.contains('艳玲')]
        
        #将其按调整排序
        original_data = original_data.sort_values(by=['调整'], ascending=False)
        
        return original_data
    
    def remove_dupl(self, ):
        
        data = self.operate_delay()
        
        print(len(data['SO号']))
        so_list = list(set(data['SO号']))#去掉重复值
        
        
        index_list = []
        for so in so_list:
            temp_index = data.loc[data['SO号'].str.contains(so)].index
            index_list.append(temp_index[0])
        
        data = data.loc[index_list, :]
        
        for col in list(self.delay_report.keys()):
            try:
                self.delay_report[col] = data[col]
                
            except:
                if col == '产品数量':                   
                    self.delay_report['产品数量'] = data['订单数量']
                elif col == '欠料数':
                    self.delay_report['欠料数'] = data['组件需求数']
                elif col == 'PO下单日期':                   
                    self.delay_report['PO下单日期'] = ''
                elif col == '物料到料信息': 
                    self.delay_report['物料到料信息'] = '延期工厂交期'
                elif col == '建议延期日期':
                    self.delay_report['建议延期日期'] = data['调整']
                else:
                     print(f'{col} not in original_data')
        
        
                    
    def save_files(self, ):
        print(self.delay_report)
                
        


if __name__ == '__main__':
    D = delayReport()
    D.operate_delay()
    D.remove_dupl()
    D.save_files()
