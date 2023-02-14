# -*- coding: utf-8 -*-
"""
Created on Sun Jul 18 16:13:58 2021

@author: 吴彦祖的laptop
"""

import pandas as pd
import numpy as np
import datetime, time
import re

def needTime(reports):
    if type(reports['DUE_DATE']) == datetime.datetime:
        if reports[0] == '成品层':
            if '2160' in reports['物料组描述(成品)'] or '2460' in reports['物料组描述(成品)']:
                reports['我司需求'] = reports['DUE_DATE'] - datetime.timedelta(days=10)
            else:
                reports['我司需求'] = reports['DUE_DATE'] - datetime.timedelta(days=6)
        else:
            reports['我司需求'] = reports['DUE_DATE'] - datetime.timedelta(days=10)

    return reports['我司需求']

def add_jit(report, jit_tuple):
    
    report['产品编码'] = str(report['产品编码'])
    
    if len(report['产品编码'])>12:
        code = str(report['产品编码'])[6:18]
    else:
        code = report['产品编码']
    
    code = int(code)
    
    if code in jit_tuple:
        report['N+2计划&JIT叫料'] = '包材JIT标识'
    
    return report

def judge(report, series_dic, person):
    '''
    number:主数据管理员的代码
    series：系列的数字
    series——dic：系列对应的产线字典
    person：记录对应关系的sheet表
    '''
    

    #欠料表中判断属于哪个部门
    number = str(report['主数据'])
    series = str(report['数字系列'])
    string = str(report['物料组描述(成品)'])
    temp_index = person.loc[person['管理员'].str.contains(number)].index
    
    #若主数据是空值，那么在contains 这个语句中，代表全选
    if len(temp_index) == 10:
        temp_index = ''
    
    if len(temp_index):
        name = person.loc[temp_index, '姓名']
        name_string = '、'.join(list(name))
    
        return name_string
    else:
        if series in series_dic:
            return series_dic[series]
        elif ('回收' in string) or ('粉筒' in string) or ('理光' in string) or ('施乐' in string) or ('OKI' in string) or ('京瓷' in string):
            return 'PD11'


def get_num(report):
    #get number from string
    report = str(report)
    if len(report) > 1:
        num = re.sub(r'\D', '', report)
    else:
        print(report)
        
    return num

def creat_dict(report, dic, string1, string2):#将透视的表变成字典
        
        dic[report[string1]] = report[string2]
        
        return dic
    
    
def read_EXCL():
    #导入数据
    addr = r'C:\Users\吴彦祖的laptop\Desktop\company\欠料表.xlsx'
    data = pd.read_excel(addr, sheet_name=None)
    
    return data
    
def cleanData(data):
    
    #获取原始缺料表
    shortageSo = data['销售订单物料缺料分析报表 (10)']
    headSo = data['SOhead']
    # #获取表头
    # headSoList = list(data['SOhead'].keys())
    # headList = list(data['head'].keys())
    
    # #获取即将创建的表的长和宽
    # totalCell = len(headSoList) * len(shortageSo)
    # #创建表
    # headSo = pd.DataFrame(np.arange(totalCell).reshape((len(shortageSo), len(headSoList))), columns=headSoList)
    
    # #英文表头切换成中文
    # shortageSo.columns = headList
    
    #将系统表的列按照我们习惯摆放
    for i in list(shortageSo.keys()):
        headSo[i] = shortageSo[i]
        
    #备注叫料和寄售信息
    consignment_dic = {}
    #如果是string类型，要变成数值类型
    consignment_dic = data['寄售'].apply(creat_dict, args=(consignment_dic, '编码', '类型'), axis=1)[0]
        
    headSo['到料信息备注'] = headSo['物料组件编码'].apply(lambda x: consignment_dic[x] if (x in consignment_dic) else '')
     
    return headSo

def operateData(headSo):
    #选出芯片
    #headSo = headSo.loc[headSo['物料组描述（物料）'].str.contains('芯片专用|全新芯片')]
    
    #将销售订单合并成我们希望的样子
    headSo['销售订单'] = headSo['销售订单'].apply(str)
    headSo['SO'] = headSo['销售订单'] + headSo['行项目']
    # headSo['SO'] = headSo['SO'].apply(lambda x:"".join(['0', x] if x[:2] == '46' else x))
    
    # #判断是成品层还是半成品
    # headSo['BOM最后一次修改日期'] = headSo['BOM最后一次修改日期'].apply(str)
    # headSo[0] = headSo['BOM最后一次修改日期'].apply(lambda x : '成品层' if x != 'NaT' else '')
    
    #整理列表index
    headSo.index = list(range(len(headSo)))
    
    #整理我司需求
    headSo['我司需求'] = headSo.apply(needTime, axis=1) 
    
    
    return headSo

def add_date(data, headSo):
    delay = data['艳玲回复交期']
    #sap_date = data['PO实际交期']
    hand_account = data['手工账']
    cor_relations = data['系列对应拉线']
    person = data['负责人']
    
    #先处理艳玲回复交期 
    #将销售订单合并成我们希望的样子
    delay['订单号'] = delay['订单号'].apply(str)
    delay['订单项目'] = delay['订单项目'].apply(str)
    delay['计划行'] = delay['计划行'].apply(str)
    delay['SO'] = delay['订单号'] + delay['订单项目'] + '_' + delay['计划行']
    delay['SO'] = delay['SO'].apply(lambda x:"".join(['0', x] if x[:2] == '46' else x))
    
    delay_dic = {}
    delay_dic = delay.apply(creat_dict, args=(delay_dic, 'SO', '核实备注'), axis=1)[0]
    print('艳玲延期', delay_dic)
    #将艳玲回复加进去
    headSo['前艳玲回复'] = headSo['SO'].apply(lambda x: delay_dic[x] if (x in delay_dic) else '')
    
    #处理实际回复交期
    #合并PO交期
    # sap_date = sap_date.drop(index=len(sap_date)-1)#最后一行为空，删除
    # sap_date['采购单号'] = sap_date['采购单号'].apply(lambda x : str(int(x)))
    # sap_date['项次'] = sap_date['项次'].apply(lambda x : str(int(x)))
    # sap_date['计划行号'] = sap_date['计划行号'].apply(lambda x : str(int(x)))
    # sap_date['PO核对'] = sap_date['采购单号'] +  + sap_date['项次'] + '_' + sap_date['计划行号']
    

    # headSo['PO单号'] = headSo['PO单号'].apply(str)
    # headSo['PO核对'] = headSo.apply(lambda x : str(int(float(x['PO单号']))) + str(x['PO行号']) if len(x['PO单号']) > 3 else '', axis=1)
    
    # date_dict = {}
    # date_dict = sap_date.apply(creat_dict, args=(date_dict, 'PO核对', '交货日期'), axis=1)[0]
    # print('PO交期', date_dict)
    
    # #将PO交期核对准
    # headSo['PO交期'] = headSo['PO核对'].apply(lambda x: date_dict[x] if (x in date_dict) else '无交期')
    
    #添加JIT标识
    jit_sign_tuple = set(list(hand_account.loc[hand_account['N+2计划&JIT叫料'] == '包材JIT标识']['物料号']))#将有jit标识的编码全部找出来
    headSo = headSo.apply(add_jit, args=(jit_sign_tuple,), axis=1)
    
    #添加计划负责人
    cor_relations['数字系列'] = cor_relations['父件物料组描述'].apply(get_num, args=())
    
    #获取数字系列对应的部门
    temp_index = cor_relations.loc[cor_relations['部门'].str.contains('PD11')].index#把PD11的去除
    cor_relations = cor_relations.drop(index=temp_index)
    
    miss = cor_relations['父件物料组描述'].isnull()#把空的值去掉
    temp_index = miss.loc[miss == True].index
    cor_relations = cor_relations.drop(index=temp_index)
    
    series_dic = {}
    series_dic = cor_relations.apply(creat_dict, args=(series_dic, '数字系列', '部门'), axis=1)[0]
    
    headSo.insert(2, '计划负责人', 0)
    
    main_data_dict = {}#主数据管理员
    main_data_dict = hand_account.apply(creat_dict, args=(main_data_dict, 'SO','主数据管理员'), axis=1)[0]
    
    headSo['主数据'] = headSo['SO'].apply(lambda x : main_data_dict[x] if x in main_data_dict else '')
    
    #把欠料表中的系列转化成数字
    headSo['数字系列'] = headSo['物料组描述(成品)'].apply(get_num, args=())
    
    #接下来就是根据两个字典，把计划负责人填上去，再根据主数据
    #通过两个元素定位计划负责人：1、主数据管理员；2、产品系列
    person['管理员'] = person['管理员'].apply(str)
    headSo['计划负责人'] = headSo.apply(judge, args=(series_dic, person), axis=1)    
    
    return headSo




def saveFiles(headSo):
    fileName = '欠料表' + time.strftime("%Y-%m-%d", time.localtime()) + '.xlsx'
    addr = 'D:\\运行结果\\' + fileName
    
    with pd.ExcelWriter(addr) as writer:
        headSo.to_excel(writer, sheet_name='欠料表', index=False)
        
def main():
    start_time = datetime.datetime.now()
    
    df = cleanData(data)
    df = operateData(df)
    df = add_date(data, df)
    saveFiles(df)
    end_time = datetime.datetime.now()
    print('程序耗时{}分'.format(((end_time-start_time).seconds / 60)))
    
if __name__ == '__main__':
    data = read_EXCL()
    main()
