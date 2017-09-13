#!/usr/bin/env python
# encoding: utf-8
"""
@version: 2.7.13
@author: tangtao
@contact: tangtao@lhtangtao.com
@description: 此处添加描述
@site: http://www.lhtangtao.com
@software: PyCharm
@file:  kaoqin

"""
import win32com.client as win32
def sort_by(num):
    """
    根据第num列来排序
    :param num:
    :return:
    """
    excel = win32.Dispatch('Excel.Application')  # 获取Excel
    wb = excel.Workbooks.Open('77.xlsx')
    ws = wb.Worksheets('Sheet1')
    ws.Range('A2:B6').Sort(Key1=ws.Range('B1'), Order1=1, Orientation=1)
    wb.Save()
    wb.Close(SaveChanges=0)



if __name__ == '__main__':
    # closesoft()

    sort_by(1)