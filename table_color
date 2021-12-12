import openpyxl as op
import pandas as pd

'''
这个还是没有封装，后面有需要的话，再考虑封装
'''


addr = 'D:\运行结果\欠料表2021-10-10.xlsx'
data = pd.read_excel(addr, sheet_name=None)

excel = op.load_workbook(addr)

s = excel['one']['A1':'AZ1']# 选中第一行

fill = op.styles.PatternFill("solid", fgColor="55ACDD")# 背景颜色
ft = op.styles.Font(color="00000000")# 字体颜色
alignment = op.styles.Alignment(horizontal='left')# 对齐方式

for i in s:
    for j in i:  
        j.fill = fill # 添加背景色
        j.font = ft # 添加字体颜色
        j.alignment = alignment # 添加对齐方式
        
        
saveAddr = 'D:\运行结果\欠料表2021-10-10.xlsx'
excel.save(saveAddr)
