#! python3
# -*- coding:utf-8 -*-
import os
import openpyxl

os.chdir('E:\\基础表\业务预算')
#修改路径
mz = []#定义空列表
mz = os.listdir()
nl = []


for i in range(len(mz)):
    wb = openpyxl.load_workbook(mz[i],data_only=True)
    nl = wb.get_sheet_names()
    if '汇总表' in nl:
        sheet = wb.get_sheet_by_name('汇总表')
        sheet.title = '汇总'
        wb.save(mz[i])