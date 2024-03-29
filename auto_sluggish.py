# -*- coding: utf-8 -*-
"""
Created on Sun Aug 15 16:33:03 2021

@author: 吴彦祖的laptop
"""

import pandas as pd
import numpy as np
import datetime
import logging
import re
import sys
sys.path.append(r'C:\Users\wanpeng.xie\Desktop\自动化')# 添加路径
from start_month_sluggish import main



class auto_sluggish:
    
    def __init__(self, ):
        addr = r'C:\Users\wanpeng.xie\Desktop\报表制作\cannt sent out\呆滞报表自动化.xlsx'
        sluggish = pd.read_excel(addr, sheet_name=None)
        self.finace_data = sluggish['财务呆滞明细']
        self.SAP_data = sluggish['SAP数据']
        self.total_sluggish = sluggish['格式']
        self.detial_sluggish = sluggish['呆滞明细']
        self.total_sluggish = sluggish['汇总']
        self.segmentation = sluggish['细分']
        self.first_time = datetime.datetime.strptime('20210802', '%Y%m%d').date()
        self.end_time = datetime.datetime.strptime('20210901', '%Y%m%d').date()
        
    def deal_SAP(self, data):
        
        data['物料工厂'] = data.apply(lambda x: str(x['料品编码']) + str(x['工厂']), axis=1)
        
        #设置时间
        
        #将新增的加进来
        data = data.loc[(data['入库日期'].dt.date > self.first_time) & (data['入库日期'].dt.date <= self.end_time)]
        
        return data
        
    def changePivot(self, report, index_name, col_name_list):
        
        '''
        此函数实现的是, 输入大类，和任意列，都给你汇总结果
        report:输入的表
        index_name:需要汇总的大类
        col_name_list：需要汇总数量的数据列
        '''
        pivotTemp = report.pivot_table(index=[index_name], values = col_name_list, aggfunc = 'sum')
        
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
    
    def creat_dict(self, report, dic, string1, string2):#将透视的表变成字典
        
        dic[report[string1]] = report[string2]
        
        return dic

    #实现自动vlookup函数
    def vlook_up(self, given_df, given_list, get_df, get_list):
        '''
        given_df:原始的表，dataframe类型
        given_list：given_list[0]是原始表的索引, given_list[1]是原始表需引用的值
        get_df：获得值的原始表，dataframe类型
        get_list：get_list[0]是给定表的索引，get_list[1]是获得值的列，如果没有会自动创建
        '''
            
        temp_dic = {}
        temp_dic = given_df.apply(self.creat_dict, args=(temp_dic, given_list[0], given_list[1]), axis=1)[0]
        
        get_df[get_list[1]] = get_df[get_list[0]].apply(lambda x : temp_dic[x] if (x in temp_dic) else -1)
        
        print('{}的类型是{}, {}的类型是{}, {}的类型是{}'.format(given_list[0], given_df[given_list[0]].dtypes, given_list[1], given_df[given_list[1]].dtypes, get_list[0], get_df[get_list[1]].dtypes))
        
        print('vlookup共匹配{}行，其中匹配成功{}行， 未成功{}行'.format(len(get_df[get_list[1]]), len(get_df[get_list[1]]) - (get_df[get_list[1]] == -1).sum(), (get_df[get_list[1]] == -1).sum()))
        
        return get_df
    
    def add_detail(self, report, ):
        
        temp_index = self.finace_data.loc[self.finace_data['物料工厂'] == report['物料工厂']].index[0]
        index = report['序号']
        finace_list = list(self.finace_data.loc[temp_index, ['物料', '物料组描述', '工厂', '库存单价', '入库日期', '物料描述']])
        
        self.total_sluggish.loc[index, '物料组描述'] = finace_list[1]
        self.total_sluggish.loc[index, '物料'] = finace_list[0]
        self.total_sluggish.loc[index, '物料工厂'] = str(int(report['物料工厂']))
        self.total_sluggish.loc[index, 0] = finace_list[3]
        self.total_sluggish.loc[index, '工厂'] = finace_list[2]
        self.total_sluggish.loc[index, '入库日期'] = finace_list[4]
        self.total_sluggish.loc[index, '物料描述'] = finace_list[5]
        self.total_sluggish.loc[index, '7月历史呆滞数量'] = report['本月结存呆滞数量-财务']
        self.total_sluggish.loc[index, '7月历史呆滞金额'] = report['本月结存呆滞数量-财务'] * finace_list[3]
            
    
    def add_detail_new(self, report, new_sluggish):
        
        
        temp_index = new_sluggish.loc[new_sluggish['物料工厂'] == report['物料工厂']].index[0]
        index = report['序号']
        
        #如果新增编码，就自动加上去
        if report['物料工厂'] in set(self.total_sluggish['物料工厂']):
            temp_index = self.total_sluggish.loc[self.total_sluggish['物料工厂'] == report['物料工厂']].index[0]
            self.total_sluggish.loc[temp_index, '7月新增数'] = report['库龄数量']
        else:
                       
            finace_list = list(new_sluggish.loc[temp_index, ['料品编码', '物料组描述', '工厂', '单位', '入库日期', '料品规格']])
            
            self.total_sluggish.loc[index, '物料组描述'] = finace_list[1]
            self.total_sluggish.loc[index, '物料'] = finace_list[0]
            self.total_sluggish.loc[index, '物料工厂'] = report['物料工厂']
            self.total_sluggish.loc[index, 0] = finace_list[3]
            self.total_sluggish.loc[index, '工厂'] = finace_list[2]
            self.total_sluggish.loc[index, '入库日期'] = finace_list[4]
            self.total_sluggish.loc[index, '物料描述'] = finace_list[5]
            self.total_sluggish.loc[index, '7月新增数'] = report['库龄数量']
            
    
    
    def test(self, ):
        
        new_sluggish = self.deal_SAP(self.SAP_data)
        
        #接下来将先将财务的处理一下
        finace_pivot = self.changePivot(self.finace_data, '物料工厂', ['本月结存呆滞数量-财务'])
        finace_pivot['序号'] = list(finace_pivot.index) #增加序号，假装在for循环
        #开始操作
        finace_pivot.apply(self.add_detail, axis=1) #i没啥用，就是记得参数如何输进去
        
        #新增的编码再加进去
        pivot_new = self.changePivot(new_sluggish, '物料工厂', ['库龄数量'])
        end_index = len(self.total_sluggish) + len(pivot_new)#循环要继续加
        pivot_new['序号'] = list(range(len(self.total_sluggish), end_index, 1))#新增编码的序号也要加进去
        
        #加入新增编码
        pivot_new.apply(self.add_detail_new, axis=1, args=(new_sluggish,))
        
        #再加入大类
        self.total_sluggish = self.vlook_up(self.segmentation, ['物料组描述', '细分'], self.total_sluggish, ['物料组描述', '细分'])
        
        #有些半成品和成品匹不到，加进去
        self.total_sluggish['细分'] = self.total_sluggish.apply(lambda x: '自制半成品' if (re.match(r'半成品(.*)系列(.*)', x['物料组描述'], re.M|re.I)) else x['细分'], axis=1)
        self.total_sluggish['细分'] = self.total_sluggish.apply(lambda x: '自制半成品' if (re.match(r'成品-(.*)', x['物料组描述'], re.M|re.I)) else x['细分'], axis=1)
        
        self.save_file()
        
    def week_deal(self, ):
        self.SAP_data['物料工厂'] = self.SAP_data.apply(lambda x : str(x['料品编码']) + str(x['工厂']), axis=1)

        #筛选出新旧编码然后透视，然后获得字典
        self.sap_data_old = self.SAP_data.loc[self.SAP_data['入库日期'].dt.date <= self.first_time]
        self.sap_data_new = self.SAP_data.loc[(self.SAP_data['入库日期'].dt.date > self.first_time) & (self.SAP_data['入库日期'].dt.date <= self.end_time)]
        old_pivot = self.changePivot(self.sap_data_old, '物料工厂', ['库龄数量'])
        new_pivot = self.changePivot(self.sap_data_new, '物料工厂', ['库龄数量'])
        old_dict = {}
        old_dict = old_pivot.apply(self.creat_dict, args=(old_dict,), axis=1)[0]
        new_dict = {}
        new_dict = new_pivot.apply(self.creat_dict, args=(new_dict,), axis=1)[0]
        
        #开始赋值
        col_name_list = ['8-18日历史呆滞数量','8-18日历史呆滞金额', '8-18日新增呆滞数量', '8-18日新增呆滞金额', '8-18日总呆滞数量','8-18日总呆滞金额','第一周消耗数量','第一周总消耗金额']
        for col in col_name_list:
            self.detial_sluggish[col] = 0#先全部赋值为0
        
        self.detial_sluggish['物料工厂'] = self.detial_sluggish['物料工厂'].apply(str)
        
        month = str(datetime.datetime.now().month)
        print(type(month))
        self.detial_sluggish[col_name_list[0]] = self.detial_sluggish['物料工厂'].apply(lambda x : old_dict[x] if (x in old_dict) else 0)
        self.detial_sluggish[col_name_list[1]] = self.detial_sluggish.apply(lambda x : x[col_name_list[0]] * x[0], axis=1)
        self.detial_sluggish[col_name_list[2]] = self.detial_sluggish['物料工厂'].apply(lambda x : new_dict[x] if (x in new_dict) else 0)
        self.detial_sluggish[col_name_list[3]] = self.detial_sluggish.apply(lambda x: x[col_name_list[2]] * x[0], axis=1)
        self.detial_sluggish[col_name_list[4]] = self.detial_sluggish.apply(lambda x : x[col_name_list[0]] + x[col_name_list[2]], axis=1)
        self.detial_sluggish[col_name_list[5]] = self.detial_sluggish.apply(lambda x : x[col_name_list[4]] * x[0], axis=1)
        
        print(type(self.detial_sluggish['{}月历史呆滞数量'.format(month)][0]), type(self.detial_sluggish['{}月新增数量'.format(month)][0]), type(self.detial_sluggish[col_name_list[4]][0]), )
        self.detial_sluggish[col_name_list[6]] = self.detial_sluggish.apply(lambda x : x['{}月历史呆滞数量'.format(month)] + x['{}月新增数量'.format(month)] - x[col_name_list[4]], axis=1)
        self.detial_sluggish[col_name_list[7]] = self.detial_sluggish.apply(lambda x :x[col_name_list[6]] * x[0], axis=1)
        
        summary = self.changePivot(self.detial_sluggish, '细分', ['第一周消耗数量', '第一周总消耗金额'])
        print(summary)
       
        
        
        self.save_seek_file(summary)
        
    def creat_dict(self, report, dic):#将透视的表变成字典
        
        dic[report['物料工厂']] = report['库龄数量']
        
        return dic
    
    def save_file(self):
        
        fileName_addr = r'C:\Users\wanpeng.xie\Desktop\报表制作\cannt sent out\财务呆滞结果1.xlsx'
        
        summary = self.changePivot(self.total_sluggish, '细分', ['7月历史呆滞数量', '7月历史呆滞金额', '7月新增数', '7月新增金额'])

        with pd.ExcelWriter(fileName_addr) as writer:
            self.total_sluggish.to_excel(writer, sheet_name='呆滞', index=None)
            summary.to_excel(writer, sheet_name='呆滞汇总', index=None)
            
    def save_seek_file(self, summary):
        fileName_addr = r'C:\Users\wanpeng.xie\Desktop\报表制作\cannt sent out\每周呆滞结果1.xlsx'
        
        summary_list = list(summary.keys())[1:] #把两个列名先弄出来
        for j in summary_list: #原表先加两列
            self.total_sluggish[j] = 0
            
        #然后填上内容
        flag = 0
        for j in summary_list:
            for i in list(self.total_sluggish['物料大类']):
                try:
                    temp_index = summary.loc[summary['细分'] == i].index    
                    self.self.total_sluggish.loc[flag, j] = (summary.loc[temp_index[0], j]/10000)#要除以10000
                except Exception as e:
                    logging.error('出现错误，错误类型:{}，报错信息:{}'.format(IndexError, e))
                flag += 1
            flag = 0 #当换一列的时候，要将索引清零
        
        with pd.ExcelWriter(fileName_addr) as writer:
            self.detial_sluggish.to_excel(writer, sheet_name='呆滞', index=None)
            #self.total_sluggish.to_excel(writer, sheet_name='呆滞汇总', index=None)
            
            
    def run(self, ):
        
        flag = input('如果是用财务的报表，请输入1；若是每周报表，请输入2：如果是每个月月底输出新增呆滞，请输入3:')
        flag = int(flag)
        
        print(flag, type(flag))
        
        if flag == 1:
            self.test()
        elif flag == 2:
            self.week_deal()
        else:
            self.start_month()
            
    def active(self, ):
        #开始的启动项目
        start_time = datetime.datetime.now()
        self.run()
        end_time = datetime.datetime.now()
        print('程序耗时{}分'.format(((end_time-start_time).seconds / 60)))
    
    def start_month(self, ):
        main()
        
if __name__ == '__main__':
    A = auto_sluggish()
    A.active()
