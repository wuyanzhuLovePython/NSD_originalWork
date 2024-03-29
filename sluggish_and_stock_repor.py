# -*- coding: utf-8 -*-
"""
Created on Sun Dec 12 19:27:10 2021

@author: 吴彦祖的laptop
"""

import re
import pandas as pd
import numpy as np
import time
from superFuction import super_function

def active():
    
    print('输入1为处理呆滞进销存报表')
    print('输入2为处理库存进销存报表')
    flag = input('请输入你需要完成的进销存报表:')
    flag = int(flag)
    
    if flag == 1:
    
        data = imput_sluggish_files()
        deal_sluggish_report(data)
    
    else:
        data = input_stock_reprot()
        deal_stock_report(data)
    
def sum(a,b):
    a = float(a)
    b = float(b)
    return a+b
    
'''
以下为处理呆滞进销存的代码
'''
    
def imput_sluggish_files():
    
    addr = r'C:\Users\wanpeng.xie\Desktop\呆滞进销存报表.xlsx'
    data = pd.read_excel(addr, sheet_name=None)
    
    return data

def deal_sluggish_report(data):
    
    total_sluggish = data['原始表']
    categories = data['类别']
    
    # 先往原始表中加入大类

    cate_list = ['大类', '细分']
    
    for i in cate_list:
        super_function.vlook_up(categories, ['物料组描述', i], total_sluggish, ['物料组描述', i])
        
    # 注意：做这一步的顺序必须在先vlookup之后再做，不然这个表就只有成品和半成品了
    # 先将大类的成品和半成品进行区分
    total_sluggish['大类'] = total_sluggish.apply(lambda x:'硒鼓半成品' if (re.match(r'半成品(.*)系列(.*)', x['物料组描述'], re.M|re.I)) else x['大类'], axis=1)
    total_sluggish['大类'] = total_sluggish.apply(lambda x:'硒鼓成品' if (re.match(r'成品(.*)', x['物料组描述'], re.M|re.I)) else x['大类'], axis=1)
    
    #再将细分的成品和半成品进行区分
    total_sluggish['细分'] = total_sluggish.apply(lambda x:'自制半成品' if (re.match(r'半成品(.*)系列(.*)', x['物料组描述'], re.M|re.I)) else x['细分'], axis=1)
    total_sluggish['细分'] = total_sluggish.apply(lambda x:'订单成品' if (re.match(r'成品(.*)', x['物料组描述'], re.M|re.I)) else x['细分'], axis=1)
    
    sluggish_summary = super_function.changePivot(total_sluggish, '大类', ['呆滞总金额-财务'])
    # 将汇总的单位从元 转化为 万元
    sluggish_summary['呆滞总金额-财务'] = sluggish_summary['呆滞总金额-财务']/10000
    
    save_sluggish_files(total_sluggish, sluggish_summary)
    

def save_sluggish_files(t_sluggish, summary):
    
    today = today = time.strftime('%Y-%m-%d', time.localtime())
    fileName_addr = r'C:\Users\wanpeng.xie\Desktop\呆滞进销存' + today + '.xlsx'
        
    with pd.ExcelWriter(fileName_addr) as writer:
        t_sluggish.to_excel(writer, sheet_name='财务呆滞报表（PMC）', index=None)
        summary.to_excel(writer, sheet_name='呆滞汇总', index=None)
        
'''
以下为处理库存进销存的代码
'''
def input_stock_reprot():
    
    addr = r'C:\Users\wanpeng.xie\Desktop\库存进销存报表.xlsx'
    data = pd.read_excel(addr, sheet_name=None)
    
    return data

def deal_stock_report(data):
    # 先加工报表，将类别和表头调整好 
    total_stock = data['原始表']
    cate_stock = data['类别']
    report_head = data['表头']
    
    # 先将表头调整一下

    for col in list(report_head.keys()):
        try:
            report_head[col] = total_stock[col]
        except:
            print('{}这个在财务的库存进销存报表中没有'.format(col))
    
    #表头这个功能到目前为止就没用了，可以赋值给总表了
    total_stock = report_head
    
    #赋值完成，这下只需要先将分类完成即可
    cate_list = ['分类2', '分类1（用于计算周转）', '细分', '细分（库存分析报表）', '品类周转报表']
    
    for cate in cate_list:
        super_function.vlook_up(cate_stock, ['物料组描述', cate], total_stock, ['物料组描述', cate])
        # 当上面匹配完了之后，成品和半成品可以直接去匹配.
        # 因为成品和半成品，经常会增加，就只能用正则表达式来处理
        if cate == '分类2' :
            total_stock[cate] = total_stock.apply(lambda x:'硒鼓半成品' if (re.match(r'半成品(.*)系列(.*)', x['物料组描述'], re.M|re.I)) else x[cate], axis=1)
            total_stock[cate] = total_stock.apply(lambda x:'硒鼓成品' if (re.match(r'成品(.*)', x['物料组描述'], re.M|re.I)) else x[cate], axis=1)
        elif cate == '分类1（用于计算周转）':
            total_stock[cate] = total_stock.apply(lambda x:'自制半成品' if (re.match(r'半成品(.*)系列(.*)', x['物料组描述'], re.M|re.I)) else x[cate], axis=1)
            total_stock[cate] = total_stock.apply(lambda x:'硒鼓成品' if (re.match(r'成品(.*)', x['物料组描述'], re.M|re.I)) else x[cate], axis=1)
        elif cate == '细分':
            total_stock[cate] = total_stock.apply(lambda x:'自制半成品' if (re.match(r'半成品(.*)系列(.*)', x['物料组描述'], re.M|re.I)) else x[cate], axis=1)
            total_stock[cate] = total_stock.apply(lambda x:'订单成品' if (re.match(r'成品(.*)', x['物料组描述'], re.M|re.I)) else x[cate], axis=1)
        elif cate == '细分（库存分析报表）':
            total_stock[cate] = total_stock.apply(lambda x:'自制半成品' if (re.match(r'半成品(.*)系列(.*)', x['物料组描述'], re.M|re.I)) else x[cate], axis=1)
            total_stock[cate] = total_stock.apply(lambda x:'成品' if (re.match(r'成品(.*)', x['物料组描述'], re.M|re.I)) else x[cate], axis=1)
        elif cate == '品类周转报表':
            total_stock[cate] = total_stock.apply(lambda x:'自制半成品' if (re.match(r'半成品(.*)系列(.*)', x['物料组描述'], re.M|re.I)) else x[cate], axis=1)
            total_stock[cate] = total_stock.apply(lambda x:'硒鼓成品' if (re.match(r'成品(.*)', x['物料组描述'], re.M|re.I)) else x[cate], axis=1)
            
    # 先做各品类月出库量，因为这个相对最简单，且只用到总表。
    stock_in_and_out = stock_in_and_out_fuction(total_stock)
    
    # 做原材料周转报表
    raw_meterial_turnover = raw_meterial_turnover_fuction(total_stock)
    
    # 做总库存周转报表
    total_turnover = total_turnover_fuction(total_stock)
    
    # 做汇总1的第一个表
    stock_total1_1 = stock_total1_1_fuction(total_stock)
    
    # 做汇总1的第二个表格
    stock_total1_2 = stock_total1_2_fuction(total_stock)
    
    # 做汇总1的第四个表格
    stock_total1_4 = stock_total1_4_fuction(total_stock)
    
    
    save_dict = {'库存进销存报表' : total_stock, '各品类月入库出库量' : stock_in_and_out, '原材料周转报表' : raw_meterial_turnover, '总周转报表' : total_turnover, '汇总1的第一报表' : stock_total1_1, '汇总1的第二报表' : stock_total1_2, '汇总1的第四报表' : stock_total1_4}
    
    save_stock_files(save_dict)
    
def stock_in_and_out_fuction(total_stock):
    
    return super_function.changePivot(total_stock, '品类周转报表', ['本期进库数量', '本期出库数量'])# 这样就完成了 各品类月出库量汇总

def raw_meterial_turnover_fuction(total_stock):
    '''
    思路：
    1、先按照物料组描述透视
    2、然后把半成品、成品等没用的删除
    3、最后将芯片和外购注塑件合并
    '''
    # 一下就是所需要的关键字，就取这几个作为关键字
    key_word = ['全新彩色碳粉', '全新充电辊', '全新出粉刀', '全新磁辊', '全新感光鼓', '全新黑色碳粉', '全新清洁刮刀', '全新显影辊', '全新送粉辊', '全新五金件', '全新辅助件', '外购注塑件', '芯片']
    
    #另起一列赋值
    total_stock['物料组描述NEW'] = total_stock['物料组描述']
    
    
    # 物料组描述中，没有'外购注塑件'、'芯片'这两个字符串，必须自己造
    total_stock['物料组描述NEW'] = total_stock['物料组描述NEW'].apply(lambda x : '外购注塑件' if ( '外购注塑件' in x ) else x)
    total_stock['物料组描述NEW'] = total_stock['物料组描述NEW'].apply(lambda x : '芯片' if ( '芯片' in x ) else x )
    
    raw_meterial_turnover = super_function.changePivot(total_stock, '物料组描述NEW', ['期初库存数量', '期初库存金额', '本期出库数量', '本期出库金额', '期末库存（财务）数量', '期末库存（财务）金额'])
    
    index_list = []
    for index in range(len(raw_meterial_turnover)):
        # 判断一下物料组描述在不在关键字里面
        group_describe = raw_meterial_turnover.loc[index, '物料组描述NEW']
        
        if group_describe in key_word:
            index_list.append(index)
    
    #将索引规范一下
    raw_meterial_turnover = raw_meterial_turnover.loc[index_list, :]
    raw_meterial_turnover.index = list(range(len(raw_meterial_turnover)))
    
    # 先全部除以10000, 即把单位转为 '万'
    raw_meterial_turnover[['期初库存数量', '期初库存金额', '本期出库数量', '本期出库金额', '期末库存（财务）数量', '期末库存（财务）金额'] ] = raw_meterial_turnover[['期初库存数量', '期初库存金额', '本期出库数量', '本期出库金额', '期末库存（财务）数量', '期末库存（财务）金额'] ]/10000
    
    # 把需要加上的内容加上。其实这个并不是很必要，但还是加上吧
    raw_meterial_turnover['周期'] = 28
    
    # 计算按照金额计算月周转天数
    raw_meterial_turnover['月周转天数（按金额核算）'] = raw_meterial_turnover.apply(lambda x : (x['周期']/(x['本期出库金额']/(x['期初库存金额'] + x['期末库存（财务）金额']) * 2)), axis=1)
    # 计算按照数量的月周转天数
    raw_meterial_turnover['月周转天数（按数量核算）'] = raw_meterial_turnover.apply(lambda x : (x['周期']/(x['本期出库数量']/(x['期初库存数量'] + x['期末库存（财务）数量']) * 2)), axis=1)
    
    return raw_meterial_turnover
    # 以上原材料周转完成了

def total_turnover_fuction(total_stock):
    '''
    做总的周转报表
    思路：
    1、用分类1这个字段进行透视
    2、总周转表中，‘硒鼓半成品’这个一行，要单独处理，因为分类2中是没有这个字段的
    '''
    
    # 透视汇总
    total_turnover = super_function.changePivot(total_stock, '分类1（用于计算周转）', ['期初库存金额', '本期出库金额', '期末库存（财务）金额'])
    # 将计算单位由‘元’转化成'万元'
    total_turnover[['期初库存金额', '本期出库金额', '期末库存（财务）金额']] = total_turnover[['期初库存金额', '本期出库金额', '期末库存（财务）金额']]/10000
    
    # 先把外购半成品和自制半成品的索引找到
    test_list = list(total_turnover['分类1（用于计算周转）'])
    purchase_halfgood_index = test_list.index('外购半成品') 
    produce_halfgood_index = test_list.index('自制半成品')
    
    # 新插入一行的做法： 切片——赋值——合并
    df_half_good = total_turnover.loc[8,:]# 对最后一行切片
    
    #赋值
    for col in list(total_turnover.keys()):
        if col == '分类1（用于计算周转）':
            df_half_good[col] = '硒鼓半成品'
        elif col == '期初库存金额':
            df_half_good[col] = '期初库存金额'
            df_half_good[col] = sum(total_turnover.loc[purchase_halfgood_index, col],  total_turnover.loc[produce_halfgood_index, col])
        elif col == '本期出库金额':
            df_half_good[col] = '本期出库金额'
            df_half_good[col] = total_turnover.loc[purchase_halfgood_index, col] + total_turnover.loc[produce_halfgood_index, col]
        elif col == '期末库存（财务）金额':
            df_half_good[col] = '期末库存（财务）金额'
            df_half_good[col] = total_turnover.loc[purchase_halfgood_index, col] + total_turnover.loc[produce_halfgood_index, col]
    
    # 合并
    total_turnover = total_turnover.append(df_half_good)
    total_turnover['周期'] = 28
    total_turnover['月周转天数'] = total_turnover.apply(lambda x : x['周期']/(x['本期出库金额']/(x['期初库存金额'] + x['期末库存（财务）金额']) * 2), axis=1)

    return total_turnover

def stock_total1_1_fuction(total_stock):
    #进销存报表中的汇总1 的第一个表格
    stock_total1_1 = pd.DataFrame(np.arange(6).reshape((1,6)), columns = ['月份', '库存总额', '周转天数', '呆滞结存金额', '呆滞计提金额', '呆滞占比'])
    
    stock_total1_1.loc[0, '库存总额'] = total_stock['期末库存（财务）金额'].sum()/10000
    
    return stock_total1_1

def stock_total1_2_fuction(total_stock):
    # 进销存报表中汇总1的第二个表格
    stock_total1_2 = super_function.changePivot(total_stock, '分类2', ['期末库存（财务）金额'])
    
    stock_total1_2['期末库存（财务）金额'] = stock_total1_2['期末库存（财务）金额']/10000
    
    return stock_total1_2

def stock_total1_4_fuction(total_stock):
    stock_total1_4 = total_stock

    #第四个表格由于没有现成的汇总，所以要从物料组描述中透视汇总。有以下两个大类需要单独处理
    #先处理外购注塑件
    stock_total1_4['物料组描述'] = stock_total1_4.apply(lambda x:'外购注塑件' if (re.match(r'外购注塑件(.*)', x['物料组描述'], re.M|re.I)) else x['物料组描述'], axis=1 )
    # 再处理芯片
    stock_total1_4['物料组描述'] = stock_total1_4.apply(lambda x:'芯片' if ( '芯片' in x['物料组描述']) else x['物料组描述'], axis=1 )
    
    #数据预处理完成后，再继续透视汇总
    
    stock_total1_4 = super_function.changePivot(stock_total1_4, '物料组描述', ['期末库存（财务）金额'])
    stock_total1_4['期末库存（财务）金额'] = stock_total1_4['期末库存（财务）金额']/10000
    
    return stock_total1_4
    
def save_stock_files(save_dict):
    
    today = today = time.strftime('%Y-%m-%d', time.localtime())
    fileName_addr = r'C:\Users\wanpeng.xie\Desktop\库存进销存' + today +'.xlsx'
        
    with pd.ExcelWriter(fileName_addr) as writer:
        for sheet_name, report in save_dict.items():
            report.to_excel(writer, sheet_name=sheet_name, index=None)



if __name__ == '__main__':
    
    active()
