
# 用于SAP的函数
def sap_output_df(self, transaction_code, number, shell):
        '''
        注意：shell的使用不要有类，不然不好引用。如果在类里面，记得也要加上类
        
        transaction_code : 事务代码
        number ： 当前shell一共有多少行
        shell ： 这个shell的调用接口
        '''
        
        number = shell.RowCount
        if transaction_code == 'coois':
            name_dict = {'销售订单':'KDAUF_AUFK', '项目行':'KDPOS_AUFK', '计划工厂':'PLWRK', '订单数量':'GAMNG', '交货数量':'GWEMG', '确认数量':'IGMNG', '基本完成时间':'GLTRP', '订单类型':'AUART', '生产管理员':'FEVOR', '系统状态':'STTXT', '创建日期':'ERDAT', '更改日期':'AEDAT', '基本开始日期':'GSTRP'}
        else:
            name_dict = {'订单':'AUFNR', '移动类型':'BWART', '抬头文本':'BKTXT', '工厂':'WERKS', '数量':'MENGE', '库位':'LGORT', '物料':'MATNR', '文本':'SGTXT', '批次':'CHARG'}
            
        df = pd.DataFrame(name_dict, pd.Index(range(1)))
        
        for i in range(number):
            for name,code in name_dict.items():
                df.loc[i, name] = shell.getCellValue(i, code)
                
        return df
