#! python3
# -*- coding:utf-8 -*-
import os
import openpyxl
os.chdir('I:\\')
a=0
wb = openpyxl.Workbook()
sheet = wb.get_active_sheet()
for rowOfCellObjects in sheet['A1':'C3']:
    for cellObj in rowOfCellObjects:
        cellObj.value=a
        a=a+1
wb.save('abc.xlsx')



end='D'+str(len(mz))
i=0
wb = openpyxl.Workbook()
sheet = wb.get_active_sheet()
for rowOfCellObjects in sheet['A1':'D4']:
    for cellObj in rowOfCellObjects:
        cellObj.value=nl[i]
        i=i+1
wb.save('abc.xlsx')
