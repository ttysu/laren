#! python3
# -*- coding:utf-8 -*-
import openpyxl
import os
os.chdir('E:\\基础表\业务预算')


wb = openpyxl.load_workbook('!!!!联通(增值税)-CBD人寿大厦-分布-业务施工.xlsx')
sheet = wb.get_sheet_by_name('汇总')
for row in sheet.iter_rows():
    for cell in row:
        if cell.value == '工程概、预算总费用':
            print(cell.row, cell.value)
        elif cell.value =='让利（建安费-材料费）*（1-折扣）':
            print(cell.row, cell.value)
        elif  cell.value == '集成费（不含材料费）':
            print(cell.row, cell.value)
print('--- END OF ROW ---')