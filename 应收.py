#! python3
# -*- coding:utf-8 -*-
import openpyxl
import os
os.chdir('E:\\基础表\业务预算')
# 修改路径
mz = []  # 定义空列表
mz = os.listdir()
nl = []
for i in range(len(mz)):
    os.chdir('E:\\基础表\业务预算')
    wb = openpyxl.load_workbook(mz[i], data_only=True)
    sheet = wb.get_sheet_by_name('汇总')
    A1 = sheet.cell(row=2, column=1).value  # 工程名
    A2 = sheet.cell(row=3, column=1).value  # 站点名
    A3 = sheet.cell(row=3, column=5).value  # 折扣
    A4 = sheet.cell(row=3, column=7).value  # 面积
    print(A2)
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == '工程概、预算总费用':
                A5 = sheet.cell(row=cell.row, column=6).value # 工程概、预算总费用
                print(cell.row, cell.value)
            elif cell.value == '让利（建安费-材料费）*（1-折扣）' or cell.value ==  '让利（建安费-材料费-消项税额）*（1-折扣）':
                A6 = sheet.cell(row=cell.row, column=6).value # 让利（建安费-材料费）*（1-折扣）
                print(cell.row, cell.value)
            elif cell.value == '集成费（不含材料费）':
                A7 = sheet.cell(row=cell.row, column=6).value #集成费（不含材料费）
                print(cell.row, cell.value)
    wb.save(mz[i])
    os.chdir('E:\\')
    wb = openpyxl.load_workbook('abc.xlsx')
    sheet = wb.get_active_sheet()
    sheet.cell(row=i + 1, column=1).value = A1
    sheet.cell(row=i + 1, column=2).value = A2
    sheet.cell(row=i + 1, column=3).value = A3
    sheet.cell(row=i + 1, column=4).value = A4
    sheet.cell(row=i + 1, column=5).value = A5
    sheet.cell(row=i + 1, column=6).value = A6
    sheet.cell(row=i + 1, column=7).value = A7
    wb.save('abc.xlsx')
