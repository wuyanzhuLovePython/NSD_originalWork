# -*- coding: utf-8 -*-
"""
Created on Sun Jan 10 21:31:36 2021

@author: wanpeng.xie
"""

import pandas as pd
import numpy as np
import win32com.client 
import re
import time
from datetime import timedelta, datetime
from superFuction import super_function

# print('------------正在导入数据-------------')
# addr = 'C:\\Users\\86151\\Desktop\\company\\assembly自动欠料表核对.xlsx'
# original = pd.read_excel(addr)
# original = original.applymap(str)
# checkChip = original.loc[original['物料组描述(物料)'].str.contains('全新芯片|芯片专用')]
# checkChip.index = range(len(checkChip))
# checkChip.insert(26, '调整', '')
# print('---------------导入成功---------------')

# SAP需要的东西
SapGuiAuto = win32com.client.GetObject("SAPGUI") 
application = SapGuiAuto.GetScriptingEngine 
connection = application.Children(0)
session = connection.Children(0)    #con、ses都抄图片拾取到的[]中的编号 session.findById("wnd[0]").resizeWorkingPane(147, 28, 0) session.findById("wnd[0]/usr/txtGS_PARTNER-DESCRIP_LONG").text = "1234" session.findById("wnd[0]/usr/txtGS_ADDRESS-NAME1").text = "456" session.findById("wnd[0]/usr/txtGS_ADDRESS-NAME2").text = "789" session.findById("wnd[0]/usr/txtGS_ADDRESS-NAME2").setFocus()

try:
    session_1 = connection.Children(1)
except:
    session.createSession()


class checkShortage:
    
    def importFiles(self, ):
        print('------------正在导入数据-------------')
        addr = r'C:\Users\wanpeng.xie\Desktop\自动欠料表核对.xlsx'
        global original, checkChip, exception_information, changed_code, hand_account
        #将需要的数据导出来
        original = pd.read_excel(addr)
        self.yesterday_short = pd.read_excel(addr, sheet_name='昨天欠料表')
        self.delay_factory = pd.read_excel(addr, sheet_name='工厂延期')
        exception_information = pd.read_excel(addr, sheet_name='例外信息')
        changed_code = pd.read_excel(addr, sheet_name='改前改后编码')
        hand_account = pd.read_excel(addr, sheet_name='手工账')
        
        original = original.applymap(str)
        checkChip = original.loc[original['物料组描述(物料)'].str.contains('全新芯片|芯片专用')]
        self.yesterday_short = self.yesterday_short.loc[self.yesterday_short['物料组描述(物料)'].str.contains('全新芯片|芯片专用')]
        
        # 先将半成品编码和描述引用过来
        checkChip['产品编码'] = checkChip['产品编码'].apply(int) # 不知道为什么，一定要转化成int才能匹配
        checkChip = super_function.vlook_up(hand_account, ['物料号', 'BOM半成品编码'], checkChip, ['产品编码', 'BOM半成品编码'])
        checkChip = super_function.vlook_up(hand_account, ['物料号', 'BOM半成品描述'], checkChip, ['产品编码', 'BOM半成品描述'])
        
        checkChip.index = range(len(checkChip))
        insertList = list(checkChip.keys())
        checkChip.insert(insertList.index('昨天到料信息'), '调整', '')#新加一列叫'调整'
        print('---------------导入成功---------------')
    
    def files(self, ):
        print(type(original.loc[0, '产品编码']))
        checkChip['DUE_DATE'] = pd.to_datetime(checkChip['DUE_DATE'])
        checkChip['计划交货日期'] = pd.to_datetime(checkChip['计划交货日期'])
        checkChip['物料到料日期'] = pd.to_datetime(checkChip['物料到料日期'])
        checkChip['PO单号'] = checkChip['PO单号'].apply(str)
        checkChip[0] = checkChip['物料组描述(成品)'].apply(lambda x : '' if ('2160' in x or '回收' in x or '2460' in x) else '成品层')
        
       
        # 未来十四天的时间结点
        fourteen_day_future = datetime.today() + timedelta(days=7)
        
        #开始自动核对欠料表
        for i in range(len(checkChip)):
            
            #先将非PO的备注
            if checkChip.loc[i, 'PO单号'][:2] == '55':
                #get today's date
                todayDate = pd.to_datetime(time.strftime("%Y-%m-%d", time.localtime()))
                VMIdelta = checkChip.loc[i, 'DUE_DATE'] - todayDate
                print(checkChip.loc[i, 'DUE_DATE'], todayDate)
                VMIdelta = int(re.search(r'(.*)days(.*)', str(VMIdelta)).group(1))
                secondDelta = checkChip.loc[i, '计划交货日期'] -  checkChip.loc[i, 'DUE_DATE']
                secondDelta = int(re.search(r'(.*) days (.*)', str(secondDelta)).group(1))
                if VMIdelta < 8 :
                    if (VMIdelta < 4 and VMIdelta >0):
                        if secondDelta == 1:
                            checkChip.loc[i, '到料信息'] = '寄售，加急调料并协调工厂交期'
                            checkChip.loc[i, '协助人'] = '谢万鹏/艳玲'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=1)
                            continue
                        elif secondDelta ==2:
                            checkChip.loc[i, '到料信息'] = '寄售，加急调料并协调工厂交期'
                            checkChip.loc[i, '协助人'] = '谢万鹏/艳玲'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=2)
                            continue
                        elif secondDelta == 3:
                            checkChip.loc[i, '到料信息'] = '寄售，加急调料并协调工厂交期'
                            checkChip.loc[i, '协助人'] = '谢万鹏/艳玲'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=3)
                            continue
                        elif secondDelta > 3:
                            checkChip.loc[i, '到料信息'] = '寄售，加急调料并协调工厂交期'
                            checkChip.loc[i, '协助人'] = '谢万鹏/艳玲'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=4)    
                            continue
                    else:
                        checkChip.loc[i, '到料信息'] = '寄售，加急调料'
                        checkChip.loc[i, '协助人'] = '谢万鹏'
                        continue
                else:
                    checkChip.loc[i, '到料信息'] = '寄售，暂时无需调料'
                    checkChip.loc[i, '协助人'] = '谢万鹏'
                    continue # it must add a 'continue' as long as finished filling in checkChip.loc[i, '到料信息']
                    
            elif checkChip.loc[i, '类别'] == '计划订单':
                
                # python 语法规定int(x)， x如果是字符串，必须为整数。否则就得先转化为float
                if int(float(checkChip.loc[i, '短缺数量'])) <= 5:
                    checkChip.loc[i, '到料信息'] = '芯片满足，跑安全库存'
                    checkChip.loc[i, '协助人'] = 'OK'
                    checkChip.loc[i, '调整'] = '不调整'
                    continue
                else:
                    checkChip.loc[i, '到料信息'] = '计划订单查看异常'
                    checkChip.loc[i, '协助人'] = 'SQE'
                    checkChip.loc[i, '调整'] = '不调整'
                    continue
          
            elif checkChip.loc[i, '类别'] == '采购申请':
                checkChip.loc[i, '到料信息'] = '采购订单安排下单'
                checkChip.loc[i, '协助人'] = '谢万鹏，下单'
                checkChip.loc[i, '调整'] = '不调整'
                continue
            if checkChip.loc[i, '类别'] == '检验批' :
                checkChip.loc[i, '到料信息'] = '在质检满足'
                checkChip.loc[i, '协助人'] = 'OK'
                checkChip.loc[i, '调整'] = '不调整'
                continue
            
            #避免因为物料到料日期为nan而报错影响程序，日期差别放这边
            firstDelta = checkChip.loc[i, 'DUE_DATE'] -  checkChip.loc[i, '物料到料日期']
            firstDelta = int(re.search(r'(.*) days (.*)', str(firstDelta)).group(1))
            secondDelta = checkChip.loc[i, '计划交货日期'] -  checkChip.loc[i, 'DUE_DATE']
            secondDelta = int(re.search(r'(.*) days (.*)', str(secondDelta)).group(1))
            
            if firstDelta < 0:
                firstDelta = 0
                
            
            if firstDelta > 6:
                if checkChip.loc[i, 0] == '成品层':
                    checkChip.loc[i, '到料信息'] = '交期满足'
                    checkChip.loc[i, '协助人'] = 'OK'
                    checkChip.loc[i, '调整'] = '不调整' 
                    continue
                else:
                    if secondDelta > 0:
                        checkChip.loc[i, '到料信息'] = '6天半成品勉强满足'
                        checkChip.loc[i, '协助人'] = 'OK'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=1)
                    else:
                        checkChip.loc[i, '到料信息'] = '6天半成品勉强满足'
                        checkChip.loc[i, '协助人'] = 'OK'
                        checkChip.loc[i, '调整'] = '不调整'
                continue
            
            if firstDelta == 6:
                if checkChip.loc[i, 0] == '成品层':
                    checkChip.loc[i, '到料信息'] = '交期满足'
                    checkChip.loc[i, '协助人'] = 'OK'
                    checkChip.loc[i, '调整'] = '不调整'
                    continue
                else:
                    if secondDelta > 0:
                        checkChip.loc[i, '到料信息'] = '协调DUE_DATE'
                        checkChip.loc[i, '协助人'] = '艳玲'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=1)
                    else:
                        checkChip.loc[i, '到料信息'] = '6天半成品勉强满足'
                        checkChip.loc[i, '协助人'] = 'OK'
                        checkChip.loc[i, '调整'] = '不调整'
                    continue
            if firstDelta == 5:
                if checkChip.loc[i, 0] == '成品层':
                    checkChip.loc[i, '到料信息'] = '交期满足'
                    checkChip.loc[i, '协助人'] = 'OK'
                    checkChip.loc[i, '调整'] = '不调整'
                else:
                    if secondDelta > 0:
                        checkChip.loc[i, '到料信息'] = '协调DUE_DATE'
                        checkChip.loc[i, '协助人'] = '艳玲'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=1)
                    else:
                        checkChip.loc[i, '到料信息'] = '加急'
                        checkChip.loc[i, '协助人'] = '加急'
                        checkChip.loc[i, '调整'] = '不调整'
                continue
            if firstDelta == 4:
                if checkChip.loc[i, 0] == '成品层':
                    checkChip.loc[i, '到料信息'] = '交期满足'
                    checkChip.loc[i, '协助人'] = 'OK'
                    checkChip.loc[i, '调整'] = '不调整'
                else:
                    if secondDelta > 2:
                        checkChip.loc[i, '到料信息'] = '协调DUE_DATE'
                        checkChip.loc[i, '协助人'] = '艳玲'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=2)
                    else:
                        if secondDelta == 1:
                            checkChip.loc[i, '到料信息'] = '协调DUE_DATE'
                            checkChip.loc[i, '协助人'] = '艳玲/加急'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=1)
                        else:
                            checkChip.loc[i, '到料信息'] = '协调EC'
                            checkChip.loc[i, '协助人'] = '庄/加急'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=1)
                continue
            if firstDelta == 3:
                if checkChip.loc[i, 0] == '成品层':
                    if secondDelta > 1:
                        checkChip.loc[i, '到料信息'] = '协调DUE_DATE'
                        checkChip.loc[i, '协助人'] = '艳玲'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=1)
                    else:
                        checkChip.loc[i, '到料信息'] = '加急'
                        checkChip.loc[i, '协助人'] = '加急'
                        checkChip.loc[i, '调整'] = '不调整'
                else:
                    if secondDelta > 3:
                        checkChip.loc[i, '到料信息'] = '协调DUE_DATE'
                        checkChip.loc[i, '协助人'] = '艳玲'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=3)
                    else:
                        if secondDelta == 2:
                            checkChip.loc[i, '到料信息'] = '协调DUE_DATE'
                            checkChip.loc[i, '协助人'] = '艳玲/加急'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=2)
                        else:
                            checkChip.loc[i, '到料信息'] = '协调EC'
                            checkChip.loc[i, '协助人'] = '庄/加急'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=2)
                continue
            if firstDelta == 2:
                if checkChip.loc[i, 0] == '成品层':
                    if secondDelta >= 2:
                        checkChip.loc[i, '到料信息'] = '协调DUE_DATE'
                        checkChip.loc[i, '协助人'] = '艳玲'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=2)
                    else:
                        if secondDelta == 1:
                            checkChip.loc[i, '到料信息'] = '协调DUE_DATE'
                            checkChip.loc[i, '协助人'] = '艳玲/加急'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=1)
                        else:
                            checkChip.loc[i, '到料信息'] = '协调EC'
                            checkChip.loc[i, '协助人'] = '庄/加急'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=1)
                else:
                    if secondDelta > 3:
                        checkChip.loc[i, '到料信息'] = '协调DUE_DATE'
                        checkChip.loc[i, '协助人'] = '艳玲/加急'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=3)
                    else:
                        checkChip.loc[i, '到料信息'] = '协调EC'
                        checkChip.loc[i, '协助人'] = '庄/加急'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=3)
                continue
            if firstDelta < 2:
                if checkChip.loc[i, 0] == '成品层' or (len(str(checkChip.loc[i, "产品编码"])) > 12 ) :
                    if secondDelta >= 4:
                        checkChip.loc[i, '到料信息'] = '协调工厂交期'
                        checkChip.loc[i, '协助人'] = '艳玲'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=4)
                        continue
                    else:
                        if secondDelta == 3:
                            checkChip.loc[i, '到料信息'] = '协调工厂交期'
                            checkChip.loc[i, '协助人'] = '艳玲/加急'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=3)
                        else:
                            checkChip.loc[i, '到料信息'] = '协调EC'
                            checkChip.loc[i, '协助人'] = '庄/加急'
                            checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=1)
                        continue
                else:
                    if secondDelta > 3:
                        checkChip.loc[i, '到料信息'] = '协调工厂交期'
                        checkChip.loc[i, '协助人'] = '艳玲/加急'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=3)
                    else:
                        checkChip.loc[i, '到料信息'] = '协调EC'
                        checkChip.loc[i, '协助人'] = '庄/加急'
                        checkChip.loc[i, '调整'] = checkChip.loc[i, 'DUE_DATE'] + timedelta(days=3)
                continue
        
        # 这些数字全部都是字符串，核对欠料表的时候不方便, 全部转化成浮点类型
        checkChip['短缺数量'] = checkChip['短缺数量'].apply(float)  
        checkChip['订单数量'] = checkChip['订单数量'].apply(float)
        checkChip['产品数量(个)'] = checkChip['产品数量(个)'].apply(float)
        
        #--------------------------------------------------------
        # 筛选出计划订单
        global plan_order
        plan_order = checkChip.loc[(checkChip['类别'] == '计划订单') & (checkChip['短缺数量'] > 5)]
        temp_index = plan_order.loc[plan_order['物料组件描述'].str.contains('硒鼓新品开发')].index
        plan_order = plan_order.drop(index = temp_index)
         
        # 场景1：放例外信息中，安排备注

        for i in ['到料信息', '协助人', '调整']:
            plan_order = super_function.vlook_up(exception_information, ['编码', i], plan_order, ['物料组件编码', i])
            
        # 场景2：维护UI   
        # 筛选的标准就是是否有改后编码，有就把改后编码都维护UI
        changed_code['旧物料编码'] = changed_code['旧物料编码'].apply(int)
        plan_order['到料信息'] = plan_order.apply(lambda x : changed_code.loc[changed_code.loc[changed_code['旧物料编码'] == int(x['物料组件编码'])].index[0], '新物料编码'] if int(x['物料组件编码']) in list(changed_code['旧物料编码']) else x['到料信息'], axis=1)
        plan_order['新物料编码描述'] = plan_order.apply(lambda x : changed_code.loc[changed_code.loc[changed_code['旧物料编码'] == int(x['物料组件编码'])].index[0], '新物料描述'] if int(x['物料组件编码']) in list(changed_code['旧物料编码']) else '', axis=1)
        plan_order['协助人'] = plan_order.apply(lambda x : '维护UI' if int(x['物料组件编码']) in list(changed_code['旧物料编码']) else x['协助人'], axis=1)
        plan_order['调整'] = plan_order.apply(lambda x : '不调整' if int(x['物料组件编码']) in list(changed_code['旧物料编码']) else x['调整'], axis=1)

        # 筛选出需要维护UI的，维护UI
        maintain_ui = plan_order.loc[plan_order['协助人'] == '维护UI']
        maintain_ui.index = list(range(len(maintain_ui)))
        print(maintain_ui)
        
        global ui_report
        ui_report = maintainUI(maintain_ui)# 这里ui表就完成了
        
        # 场景3：OM的主替料是否错误
        # 1、先判断是否打了Z3
        # 2、再判断在成品层还是半成品层
        # 3、调用函数返回BOM数据。
        # 4、判断是否主替料错误：如果BOM里面有两个芯片编码，且打了Z3的使用比例为100.就是主替料错误，单独备注好并整理出来。
        # 5、这里注意一点，维护UI和检查BOM的主替料部分会重叠的，最后欠料表还是以检查BOM的备注为主，因为将BOM调整准确才是最重要。
        z3_index = plan_order.loc[plan_order['物料状态标识'] == 'Z3'].index
        for index in z3_index:
            
            if '硒鼓新品开发' in plan_order.loc[index, '物料组件描述']:
                plan_order.loc[index, '到料信息'] = '新品芯片，暂时不处理'
                plan_order.loc[index, '协助人'] = '产品部'
                # 芯片开发芯片不用管
                continue
            
            code = plan_order.loc[index, '物料组件编码']
            material_cate = plan_order.loc[index, '物料组描述(物料)']
            factory = plan_order.loc[index, '工厂']
            # 芯片在半成品和成品层都要，要区分一下
            if plan_order.loc[index, '层次'] == 1:
                upper_code = plan_order.loc[index, '产品编码']
            else:
                upper_code = plan_order.loc[index, '产品编码']
            
            z3_dict = check_BOM(code, upper_code, material_cate, factory)
            # 把内容加进去
            for i in ['到料信息', '协助人', '调整']:
                plan_order.loc[index, i] = z3_dict[i]  
                
                
        # 场景四：当前工厂交期在30天之后，计划订单不会转成PR。
        # 筛选出有异常的计划订单
        thirty_plan_order = plan_order.loc[plan_order['到料信息'] == -1]
        
        # 未来十四天的时间结点
        thirty_day_future = datetime.today() + timedelta(days=30)
        # 今天和工厂交期相隔天数
        #dueDate_reduc_today = (four_day_late - datetime.today()).days
        
        thirty_plan_order['到料信息'] = thirty_plan_order['DUE_DATE'].apply(lambda x : '工厂交期在30天外，未跑PR' if x >= thirty_day_future else '距离工厂交期还有{}天，找IT确认'.format((x - datetime.today()).days))
        thirty_plan_order['协助人'] = thirty_plan_order['DUE_DATE'].apply(lambda x : 'OK' if x >= thirty_day_future else 'IT/谢万鹏'.format((x - datetime.today()).days))
        
        #将内容赋值过去
        for i in thirty_plan_order.index:
            plan_order.loc[i, '到料信息'] = thirty_plan_order.loc[i, '到料信息']
            plan_order.loc[i, '协助人'] = thirty_plan_order.loc[i, '协助人']
            
        #场景4已经完成
        #---------------------------------------------------------------------------------------
        
        check_index = plan_order.index
        # 然后将核对的到料信息赋值过去
        for i in check_index:
            checkChip.loc[i, '到料信息'] = plan_order.loc[i, '到料信息']
            checkChip.loc[i, '协助人'] = plan_order.loc[i, '协助人']
            checkChip.loc[i, '调整'] = plan_order.loc[i, '调整']
            
            
        # -------------------------------------------------------
        # 找出昨天不欠今天欠的SO
        # 合并列，找出索引
        super_function.merge_only_one(self.yesterday_short, ['销售订单', '行项目', '物料组件编码'])
        super_function.merge_only_one(checkChip, ['销售订单', '行项目', '物料组件编码'])
        
        # 开始vlookup看昨天不欠今天欠的
        super_function.vlook_up(self.yesterday_short, ['核对', '到料信息'], checkChip, ['核对', '昨天到料信息'])
        temp_index = checkChip.loc[checkChip['昨天到料信息'] == -1].index
        checkChip.loc[temp_index, '昨天到料信息'] = '昨天不欠今天欠'
        
        # 挑选出来未来十四天，昨天不欠今天欠的
        sap_check = checkChip.loc[(checkChip['DUE_DATE'] < fourteen_day_future)]
        
        # 但是如果原本就不欠料的，就不用管了。所以把需要延期EC的挑出来就可以了
        check_index = sap_check.loc[(sap_check['协助人'].str.contains('庄')) | (sap_check['类别'] == '采购申请')].index
        
        # 然后一行一行地核对，先看到底欠不欠。
        # 如果真的欠，再通过mb52 和coois 看到底是为什么欠
        sap_check = coois(checkChip, check_index)
        
        # 然后将核对的到料信息赋值过去
        # 因为sap_check只有十四天的数据，不能将其直接赋值给checkChip
        for i in check_index:
            checkChip.loc[i, '到料信息'] = sap_check.loc[i, '到料信息']
            
        
        self.delay_EC = operate_delay(checkChip, self.delay_factory, '艳玲|庄')
        #delay_factory = operate_delay(checkChip, self.delay_factory, '艳玲')
        

                
    def saveFiles(self,):
        fileName = '核对欠料' + time.strftime("%Y-%m-%d", time.localtime()) + '.xlsx'
        addr = r'C:\Users\wanpeng.xie\Desktop\核对欠料表更改\\' + fileName
        with pd.ExcelWriter(addr) as writer:
            checkChip.to_excel(writer, sheet_name='核对欠料', index=None)
            self.yesterday_short.to_excel(writer, sheet_name='昨天欠料情况', index=None)
            self.delay_EC.to_excel(writer, sheet_name='延期', index=None)
            ui_report.to_excel(writer, sheet_name='维护UI', index=None)

                                 
                    
def coois(data, index):
    
    '''
    data : 输入的dataframe 的数据
    index : 输入表的索引，循环就按照这个索引
    返回的数据就是重新核对过的
    '''
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "coois"
    session.findById("wnd[0]").sendVKey(0)
    
    for i in index:
        so = data.loc[i, '销售订单']
        so_rank = data.loc[i, '行项目'][:-2]
        category = data.loc[i, '类别']
        code_material = data.loc[i, '物料组件编码']
        product_quantity = data.loc[i, '产品数量(个)']
        
        # 对于有例外信息的芯片，一起跳过
        if code_material in list(exception_information['编码']):
            continue
        
        print('开始循环',so, so_rank, category, code_material, i)
        
        if category == '采购订单' or category == "采购申请":
            #把该输入的参数输入进去
            session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_KDAUF-LOW").text = so
            session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_KDPOS-LOW").text = so_rank
            session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_COMPO-LOW").text = code_material
            session.findById("wnd[0]").sendVKey(8) #进去
          
                
            #利用报错信息判断SAP自动化是否进入窗口
            flag = 0
            try:
                id_beach = "wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell"
                shell = session.findById(id_beach)#如果没有表格，这一行就会报错
            except:
                flag = 1
            
            #确认是否进入下一个界面后
            if flag :
                data.loc[i, '到料信息'] = '无工单'
                #若没有信息,就回车键回来，然后进行下一个循环
                print(so, so_rank, '无工单', code_material, i)
                session.findById("wnd[1]").sendVKey(0)
                continue
            else:
                #若有信息，就继续判断，看是否欠料
                number = shell.RowCount#看一下有几行
                df = sap_output_df('coois', number, shell)
                
                if number == 2:
                    # 如果只有一行已经关单
                    cell_state = df.loc[0, '系统状态']# 返回工单状态
                    complet_amount = int(df.loc[0, '确认数量'])
                    if 'TECO' in cell_state and complet_amount > 0:
                        data.loc[i, '到料信息'] = '已经关单，且完成数量为{}'.format(complet_amount)
                        print(so, so_rank, '已经关单，且完成数量为{}'.format(complet_amount), i)
                        session.findById("wnd[0]/tbar[0]/btn[15]").press()# 按一次返回键就好
                        continue
                    elif'TECO' in cell_state and complet_amount == 0:
                        data.loc[i, '到料信息'] = '已经关单，但是没有做'
                        print(so, so_rank, '已经关单，但是没有做', i)
                        session.findById("wnd[0]/tbar[0]/btn[15]").press()# 按一次返回键就好
                        continue
                    
                    #如果只有一行，直接检查物料可用性
                    shell.setCurrentCell(0, "AUFNR")#选择订单行
                    session.findById("wnd[0]").sendVKey(18)#点‘笔’
                    try:
                        session.findById("wnd[0]").sendVKey(28)#检查物料可用性
                    except:
                        session.findById("wnd[1]").sendVKey(0)
                        coois_back()
                    
                    code = session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]").text #检查物料可用性获取的文本
                    
                    
                    if '均可用' in code:
                        data.loc[i, '到料信息'] = '不欠料'
                        print(so, so_rank, '均可用', i)
                        coois_back()
                        continue

                    else:
                        # 当关单的时候，这个按钮是点不了了
                        session.findById("wnd[1]/usr/btnDY_VAROPTION3").press()#点“物料”进入界面
                        #coois_back() 这时候不应该点返回
                        code_list = []
                        

                        for j in range(25):
                            code = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0200/ctxtRESBD-MATNR[1,{}]".format(j)).text
                            if len(code) == 0:
                                print('完成，一共有',j,'行')
                                break
                            code_list.append(code)
    
                        code_material = str(code_material) #因为SAP的数字都是字符格式
                        if code_material in code_list:
                            data.loc[i, '到料信息'] = '欠芯片'
                            print(so, so_rank, '欠芯片', code_material, i)
                            coois_back()
                        else:
                            data.loc[i, '到料信息'] = '不欠芯片'
                            print(so, so_rank, '不欠芯片', code_material, i)
                            session.findById("wnd[0]/tbar[0]/btn[15]").press()# 不知道为什么，当不欠芯片的时候，这个一键回原始界面了
                            # session.findById("wnd[0]/tbar[0]/okcd").text = "coois"
                            # session.findById("wnd[0]").sendVKey(0)
                            session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                            session.findById("wnd[0]/tbar[0]/btn[15]").press()
                        continue
                    
                else:
                    cell_product_quantity = shell.getCellValue(number-1, 'GAMNG')
                    cell_product_quantity = int(re.sub(r'\D', '', cell_product_quantity)) # SAP的千位数有'2,000' 这样的傻逼数字，正则替换
                    factory_list = []
                    
                    if int(product_quantity) > cell_product_quantity:
                        data.loc[i, '到料信息'] = '工单重复'
                        print(so, so_rank, '工单重复', code_material, i)
                    else:
                        data.loc[i, '到料信息'] = '订单拆单，需核查'
                        print(so, so_rank, '订单拆单，需核查', code_material, i)
                        
                    session.findById("wnd[0]/tbar[0]/btn[15]").press()
                    continue

        else:
            data.loc[i, '到料信息'] = '目前还核对不出来'
            
    return data



def coois_back():
    '''
    就是碰到无工单的情况，连续按两次返回键
    若是碰到弹窗，直接按下确定键
    '''
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    except:
        print('无弹窗，不需要')
    session.findById("wnd[0]/tbar[0]/btn[15]").press()

def coois_back1():
    '''
    zpp026在第二个窗口，session要改一下
    '''
    session_1.findById("wnd[0]/tbar[0]/btn[15]").press()
    try:
        session_1.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    except:
        print('无弹窗，不需要')
    #session_1.findById("wnd[0]/tbar[0]/btn[15]").press()
    
# 用于创建表格
def sap_output_df(transaction_code, number, shell):
    '''
    注意：shell的使用不要有类，不然不好引用。如果在类里面，记得也要加上类
    
    transaction_code : 事务代码
    number ： 当前shell一共有多少行
    shell ： 这个shell的调用接口
    '''
    
    number = shell.RowCount
    if transaction_code == 'coois':
        name_dict = {'销售订单':'KDAUF_AUFK', '项目行':'KDPOS_AUFK', '计划工厂':'PLWRK', '订单数量':'GAMNG', '交货数量':'GWEMG', '确认数量':'IGMNG', '基本完成时间':'GLTRP', '订单类型':'AUART', '生产管理员':'FEVOR', '系统状态':'STTXT', '创建日期':'ERDAT', '更改日期':'AEDAT', '基本开始日期':'GSTRP'}
    elif transaction_code == 'mb51':
        name_dict = {'订单':'AUFNR', '移动类型':'BWART', '抬头文本':'BKTXT', '工厂':'WERKS', '数量':'MENGE', '库位':'LGORT', '物料':'MATNR', '文本':'SGTXT', '批次':'CHARG'}
    elif transaction_code == 'zpp026':    
        name_dict = {'父件编码':'MATNR', '子件编码':'IDNRK', '子件描述':'MAKTX', '替代组别':'ALPGR', '优先级':'ALPRF', '使用概率':'EWAHR', '父件描述':'PMAKTX', '子件物料组描述':'WGBEZ', '子件物料组':'MATKL'}
    elif transaction_code == 'zmm003':
        name_dict = {'物料编码':'MATNR', '物料描述':'MAKTX', '工厂':'WERKS', '安全库存':'EISBE' ,'特定工厂状态':'MMSTA'}
    
    df = pd.DataFrame(name_dict, pd.Index(range(1)))
    
    for i in range(number):
        for name,code in name_dict.items():
            df.loc[i, name] = shell.getCellValue(i, code)
            
    return df

def maintainUI(maintain_ui):
    '''
    通过欠料表维护UI
    1、maintain_ui其实就是欠料表。
    2、欠料表是没有‘BOM半成品编码’，‘BOM半成品描述’的字段。如果要用此函数，需手工加上
    3、‘到料信息’字段填的是最新的编码
    ''' 
    ui_report = pd.DataFrame(np.arange(18).reshape((1,18)), columns=['序号', '日期', '是否改BOM', '销售订单', '项目', '客户代码', '成品编码', '成品描述', '总单数量', '更改数量', '更改方式', '更改前编码', '名称规格', '更改后编码', '名称规格.1', '单位', '修改原因', '备注'])
    for i in range(len(maintain_ui)):

        ui_report.loc[i, '序号'] = int(i)
        ui_report.loc[i, '日期'] = time.strftime("%Y-%m-%d", time.localtime())
        ui_report.loc[i, '是否改BOM'] = '否'
        ui_report.loc[i, '销售订单'] = maintain_ui.loc[i, '销售订单']
        ui_report.loc[i, '项目'] = re.search(r'(.*)_1', maintain_ui['行项目'][i], re.M|re.I).group(1)
        ui_report.loc[i, '客户代码'] = maintain_ui.loc[i, '客户']

        # 成品层和半成品层不一样
        if maintain_ui.loc[i, '层次'] == "1":
            ui_report.loc[i, '成品编码'] = maintain_ui.loc[i, '产品编码']
            ui_report.loc[i, '成品描述'] = maintain_ui.loc[i, '物料组描述(成品)']
        else:
            ui_report.loc[i, '成品编码'] = str(maintain_ui.loc[i, 'BOM半成品编码'])
            ui_report.loc[i, '成品描述'] = str(maintain_ui.loc[i, 'BOM半成品描述'])

        ui_report.loc[i, '总单数量'] = maintain_ui.loc[i, '订单数量']
        ui_report.loc[i, '更改数量'] = maintain_ui.loc[i, '订单数量']
        ui_report.loc[i, '更改方式'] = '替换'
        ui_report.loc[i, '更改前编码'] = maintain_ui.loc[i, '物料组件编码']
        ui_report.loc[i, '名称规格'] = maintain_ui.loc[i, '物料组件描述']
        ui_report.loc[i, '更改后编码'] = str(maintain_ui.loc[i, '到料信息'])
        ui_report.loc[i, '名称规格.1'] = str(maintain_ui.loc[i, '新物料编码描述'])
        ui_report.loc[i, '单位'] = 'PCS'
        ui_report.loc[i, '修改原因'] = 'C'
        ui_report.loc[i, '备注'] = '芯片换版，用改后编码'
    
    return ui_report

def check_BOM(code, upper_code, material_descr, factory):
    '''
    功能：检查主替料是否建错
    code : 这个是物料编码，注意这个编码一定不能是字符串
    upper_code ：这个是物料上层编码，用于在zpp026输入的
    material_descr ：这个输入的是子键物料组描述。芯片就分全新芯片和芯片专用，留这个参数是为以后全部物料都进行分析时使用
    '''
    print('现在开始检查BOM，当前编码为{}, 上层编码为{}, 物料组描述为{}, 所在工厂为{}'.format(code, upper_code, material_descr, factory))
    # 以下是sap自动查询zpp026的代码，后续再加入
    session_1.findById("wnd[0]/tbar[0]/okcd").text = "zpp026"
    session_1.findById("wnd[0]/tbar[0]/btn[0]").press() 
    
    session_1.findById("wnd[0]/usr/txtP_STLAL").text = "1"
    session_1.findById("wnd[0]/usr/ctxtP_MATNR-LOW").text = upper_code
    session_1.findById("wnd[0]/usr/ctxtP_WERKS-LOW").text = factory
    
    zpp026_tree = 'wnd[0]/usr/cntlGRID1/shellcont/shell'
    session_1.findById("wnd[0]").sendVKey(8)
    shell = session_1.findById(zpp026_tree)
    number = shell.RowCount
        
    zpp026 = sap_output_df('zpp026', number, shell)
    coois_back1()# 获取内容后就可以返回了
    
    # zpp026就是返回的内容，开始进行判断
    # 1、先将需要的物料描述筛选出来
    zpp026_1 = zpp026.loc[zpp026['子件物料组描述'] == material_descr]
    
    out_dict = {}
    # 2、找出这个编码，先确认是否有主替料
    if len(zpp026_1) <= 1:
        out_dict = {'到料信息':'BOM中无主替料，维护UI并通知产品部', '协助人':'产品部', '调整':''}
    else:
        # 3、有主替料，确认是否建错主替料
        try:
            temp_index = zpp026_1.loc[zpp026['子件编码'] == code].index[0]
            
            if zpp026_1.loc[temp_index, '使用概率'] == '100':
                out_dict = {'到料信息':'主替料错误，打了Z3的物料设置为主料', '协助人':'产品部', '调整':''}

            else:
                out_dict = {'到料信息':'主替料设置没问题，核查看是否有其他问题', '协助人':'产品部', '调整':''}
                
        except :
            #print('当前筛选出来的索引长度为{}'.format(len(temp_index)))
            out_dict = {'到料信息':'未找到索引', '协助人':'', '调整':''}

    print(out_dict)
    return out_dict 
    
def operate_delay(original_data, delay_report, name):
    
    '''
    original_data ：核对完成后的欠料表
    delay_report : 工厂延期的格式，即表头
    name ： 延期类型
    '''
    
    print('-------------开始延期{}-----------------'.format(name))
    
   
    original_data['协助人'] = original_data['协助人'].apply(str)# 只有变成string才能筛选，而且这个还能对付空值
    original_data = original_data.loc[original_data['协助人'].str.contains(name)]# 将需要延期工厂交期的找出来
    
    #将其按调整排序
    try:
        original_data = original_data.sort_values(by=['调整'], ascending=False)
    except:
        print('错误！延期日期不完整！')
        
    
    data = original_data# 其实这一个赋值是不用的，只是之前写的代码有date这个变量，懒得再替换了
    
    try:
        print(len(data['SO']))
    except:
        data['SO'] = data.apply(lambda x : x['销售订单'] + x['行项目'])
    
    so_list = list(set(data['SO']))#去掉重复值
    index_list = []
    
    # 做法就是将重复的so选出来，然后只要第一个so的index
    # 这样就把重复值剔除了
    for so in so_list:
        temp_index = data.loc[data['SO'].str.contains(so)].index
        index_list.append(temp_index[0])
    
    data = data.loc[index_list, :]
    
    print(data)
    
    for col in list(delay_report.keys()):
        try:
            delay_report[col] = data[col]
            
        except:
            if col == '产品数量':                   
                delay_report['产品数量'] = data['产品数量(个)']
            elif col == '欠料数':
                delay_report['欠料数'] = data['短缺数量']
            elif col == 'PO下单日期':                   
                delay_report['PO下单日期'] = data['PO创建日期']
            elif col == '物料到料信息': 
                delay_report['物料到料信息'] = data['到料信息']
            elif col == '建议延期日期':
                delay_report['建议延期日期'] = data['调整']
            elif col == '物料组':
                delay_report['物料组'] = '芯片'
            elif col == '延期类型':
                delay_report['延期类型'] = data['协助人']
            else:
                 print(f'{col} not in original_data')
    
    delay_report['申请日期'] = delay_report['申请日期'].apply(lambda x : datetime.today())  
    
    print(delay_report)
    
    return delay_report


    
    
if __name__ == '__main__':    
    pd.set_option('mode.chained_assignment', None)#取消警告
    C = checkShortage()
    C.importFiles()
    print('-------------------正在核对欠料表，请等待--------------------')
    C.files()
    C.saveFiles()
    print('-------------------已经核对完成------------------------')

    
    
