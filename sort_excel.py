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
import os


def sort_by(num, file_path):
    """
    根据第num列来排序
    :param num:
    :return:
    """
    excel = win32.Dispatch('Excel.Application')  # 获取Excel
    wb = excel.Workbooks.Open(file_path)
    ws = wb.Worksheets('Sheet1')
    ws.Range('A1:M321').Sort(Key1=ws.Range(num), Order1=1, Orientation=1)  # 使用win32的api进行排序
    wb.Save()
    wb.Close(SaveChanges=0)


if __name__ == '__main__':
    # closesoft()
    total_base_dir = (os.path.dirname(__file__) + '/all.xlsx').replace('/', "\\")
    sort_by('C1', total_base_dir)
