# -*- coding: utf-8 -*-
"""
Created on Sun Dec 12 15:14:46 2021

@author: 吴彦祖的laptop
"""

import pandas as pd
import numpy as np

def creat_dict(report, dic, string1, string2):#将透视的表变成字典
    
    dic[report[string1]] = report[string2]
    
    return dic

class super_function:
    
    def changePivot(report, index_name, col_name_list):
        '''
        此函数实现的是, 输入大类，和任意列，都给你汇总结果
        report:输入的表
        index_name:需要汇总的大类
        col_name_list：需要汇总数量的数据列.(注意，这个是输入列表！！)
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
    
        
    #实现自动vlookup函数
    def vlook_up(given_df, given_list, get_df, get_list):
        '''
        given_df:原始的表，dataframe类型
        given_list：given_list[0]是原始表的索引, given_list[1]是原始表需引用的值
        get_df：获得值的原始表，dataframe类型
        get_list：get_list[0]是给定表的索引，get_list[1]是获得值的列，如果没有会自动创建
        '''
            
        temp_dic = {}
        temp_dic = given_df.apply(creat_dict, args=(temp_dic, given_list[0], given_list[1]), axis=1)[0]
        
        get_df[get_list[1]] = get_df[get_list[0]].apply(lambda x : temp_dic[x] if (x in temp_dic) else -1)
        
        print('{}的类型是{}, {}的类型是{}, {}的类型是{}'.format(given_list[0], given_df[given_list[0]].dtypes, given_list[1], given_df[given_list[1]].dtypes, get_list[0], get_df[get_list[1]].dtypes))
        
        #print('vlookup共匹配{}行，其中匹配成功{}行， 未成功{}行'.format(len(get_df[get_list[1]]), len(get_df[get_list[1]]) - (get_df[get_list[1]] == -1).sum(), (get_df[get_list[1]] == -1).sum()))
        
        return get_df
    
    #自动取唯一值的函数, 其实就是新获取一列，然后这一列是一个唯一值
    def merge_only_one(data, col_name_list):
        '''
        data:原始的表，dataframe类型
        col_name_list：组成唯一值的列名
        '''
    
        col_string_list = []
        #判断每一列是否是float类型的
        for col in col_name_list:
            #拿一个来判断就好了
            #不过要先找出非缺失值index
            not_null_index = data.loc[data[col].isnull() == False].index[0]
            if '.' in str(data.loc[not_null_index, col]): #如果是float值，肯定有小数点，就要将其转化成str
                data[col] = data[col].fillna(0).astype('int64')
                data[col] = data[col].fillna(0).astype('str')
            else:
                data[col] = data[col].fillna(0).astype('str')
                
            col_string_list.append('x["{}"]'.format(col))#将列名按照元素格式化
            
        code = '+'.join(col_string_list)  #将代码变成字符串 
        
        final_code = 'data.apply(lambda x : {} ,axis=1)'.format(code)
       
        data["核对"] =  eval(final_code) #eval()这个函数不能执行 = 这样的赋值语句，因此要分开
        
        return data
    
    
        

        
