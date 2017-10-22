#! python3
#-*- coding: utf-8 -*
import os
import openpyxl

os.chdir('I:\\修改增值税预算\设计\预算')
#修改路径
mz = []#定义空列表
mz = os.listdir()
nl = []


for i in range(len(mz)-1):
    wb = openpyxl.load_workbook(mz[i],data_only=True)
    sheet = wb.get_sheet_by_name('汇总')
    A1 = sheet.cell(row=3,column=1).value
    nl.append(A1)
    wb.save(mz[i])
print(nl)

