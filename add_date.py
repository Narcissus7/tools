#  在Excel最后一列加上日期、删除前两行、导入数据库

import os
import xlrd
import xlutils.copy
import re
import win32com.client  # 这里用到win32com.client，需要安装pywin32模块
import pymysql
import pandas as pd
import sys
from sqlalchemy import create_engine


# 文件路径
path = 'E:\\旧每日数据\\12月数据\\20171115/'


# 添加日期、商品编码函数
def add_date():
    for f in os.listdir(path):  # 要处理的excel文件路径
        # print("file:", f )
        try:
            print("file:", f)
            name = re.match('.*.xls', str(f)).group()
            # print(name)
            rb = xlrd.open_workbook(path+name)  # 打开excel
            sheet = rb.sheet_by_index(0)  # 获得sheet
            # date = sheet.cell(1,0).value[5:13]  # 获取第二行第一列的内容
            date = path[-9:-1]
            # code = name[12:20]
            # code = name[8:16]
            code = re.search("[0-9]{8}", name).group()
            print(code)
            print(date)
            wb = xlutils.copy.copy(rb)
            ws = wb.get_sheet(0)
            # print(sheet.nrows)
            ws.write(2, 13, '商品编码')
            ws.write(2, 14, '日期')  # 在第三行14列写入日期

            for rows in range(3, sheet.nrows):
                ws.write(rows, 14, date)  # 在第三行到最后一行14列写入获取的时间
            for rows in range(3, sheet.nrows):
                ws.write(rows, 13, code)  # 在第三行到最后一行13列写入获取的编码
            wb.save(path + f)
        except AttributeError:
            print("file:")


def del_row():
    xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
    for f in os.listdir(path):
        try:
            name = re.match('.*.xls', str(f)).group()
            print("file:", name)
            xlBook = xlApp.Workbooks.Open(path + name)
            xlSht = xlBook.Worksheets('客户日订货情况表')  # 要处理的excel页，默认第一页是‘sheet1’
            for i in range(1, 2):
                xlSht.Rows(i).Delete()
            xlBook.Close(SaveChanges=1)  # 完成 关闭保存文件
        except AttributeError:
            continue
    del xlApp


def to_mysql():

    for f in os.listdir(path):
        # print("file:", f)
        try:
            name = re.match('.*.xls', str(f)).group()
            df = pd.read_excel(path + name)
            print(name)
            # df = df.ix[1:, [1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]]  # 行，列
            df = df.ix[1:, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]]
            print(df.head())
            yconnect = create_engine('mysql+pymysql://root:root@192.168.31.130:3306/pos?charset=utf8')
            pd.io.sql.to_sql(df, 'dec', yconnect, if_exists='append',index=None)
        except AttributeError:
            print("file:")


add_date()
del_row()
del_row()
to_mysql()