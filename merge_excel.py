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

# 将多个Excel文件合并成一个
import xlrd
import xlsxwriter

datavalue = []
# 打开一个excel文件
def open_xls(file):
    fh = xlrd.open_workbook(file)
    return fh


# 获取excel中所有的sheet表
def get_sheet(fh):
    return fh.sheets()


# 获取sheet表的行数
def get_nrows(fh, sheet):
    table = fh.sheets()[sheet]
    return table.nrows


# 读取文件内容并返回行内容
def get_Filect(file, shnum):
    fh = open_xls(file)
    table = fh.sheets()[shnum]
    num = table.nrows
    for row in range(1, num):  # 跳过第一行的数据
        rdata = table.row_values(row)
        datavalue.append(rdata)
    return datavalue


def get_shnum(fh):
    """
    获取excel中表的个数
    :param fh:
    :return:
    """
    x = 0
    sh = get_sheet(fh)
    for sheet in sh:
        x += 1
    return x


def merge(excel_list, merge_name):
    """
    此处输入你要合并的excel文件名列表以及合并后的excel的名字
    :param excel_list:
    :param merge_name:
    :return:
    """
    wb1 = xlsxwriter.Workbook(merge_name)
    # 创建一个sheet工作对象
    ws = wb1.add_worksheet()
    for fl in excel_list:
        fh = open_xls(fl)
        x = get_shnum(fh)
        for shnum in range(x):
            print("正在读取文件：" + str(fl) + "的第" + str(shnum) + "个sheet表的内容...")
            rvalue = get_Filect(fl, shnum)
    for a in range(len(rvalue)):
        for b in range(len(rvalue[a])):
            c = rvalue[a][b]
            ws.write(a, b, c)
    wb1.close()
    print("文件合并完成")


if __name__ == '__main__':
    allxls = ['1.xlsx', '2.xlsx', '3.xlsx', '4.xlsx', '5.xlsx', '6.xlsx', '7.xlsx', '8.xlsx', '9.xlsx',
              '10.xlsx', '11.xlsx', '101.xlsx', '102.xlsx', '103.xlsx', '104.xlsx', '105.xlsx', '106.xlsx',
              '107.xlsx']  # 此处输入你要合并的excel文件的文件名
    merge(allxls, 'all.xlsx')

