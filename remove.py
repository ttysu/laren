#! python3
# -*- coding:utf-8 -*-
import os
import glob
import openpyxl
wb = openpyxl.load_workbook('I:\\设计部进度表.xlsx', data_only=True)
sheet = wb.get_sheet_by_name('增值税')
os.chdir('E:\\电信')
for i in range(52):
    A1 = sheet.cell(row=i+3, column=2).value
    print(A1)
    chazhao = '*'+A1+'*'
    tt=glob.glob(chazhao)
    print(tt)
    if tt !=None:
        for j in range(len(tt)):
            os.remove(tt[j])
    tt=[]
