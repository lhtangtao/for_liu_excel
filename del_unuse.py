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

import xlrd
import xlsxwriter

from merge_excel import open_xls, get_shnum, get_Filect


def read_info(src_excel):
    """
    从源excel中读取要过滤的信息
    :param src_excel:
    :return:
    """
    datavalue = []
    excel_file = xlrd.open_workbook(src_excel)
    table = excel_file.sheets()[0]
    num = table.nrows
    for row in range(1, num):  # 跳过第一行的数据
        rdata = table.row_values(row)
        datavalue.append(rdata)
    return datavalue  # 返回一个二维数组


def del_useless(data_list):
    """
    输入一个二维数组，删除没用的信息 最后返回一个数组
    :param data_list:
    :return:
    """
    len_of_excel = len(data_list)  # excel的行数
    number_to_del = []
    for i in range(len_of_excel - 1):
        if data_list[i][2] == data_list[i + 1][2]:
            number_to_del.append(i)
    print number_to_del
    for i in range(len(number_to_del)):
        del data_list[number_to_del[i] - i]  # 删除重复的考勤记录
    return data_list


def get_useless(data_list):
    """
    输入一个二维数组，筛选出重复的信息 最后返回一个数组
    :param data_list:
    :return:
    """
    len_of_excel = len(data_list)  # excel的行数
    number_to_get = []
    for i in range(len_of_excel - 1):
        if data_list[i][2] == data_list[i + 1][2]:
            number_to_get.append(i)
    print u'以下数据是出现重复的数据，请自行加1'
    print number_to_get
    get_duplication = []
    for x in range(len(number_to_get)):
        get_duplication.append(data_list[number_to_get[x]])
    return get_duplication


def to_new_excel(filename, date_src):
    """
    把资源写入到excel中
    :param filename:
    :param date_src:
    :return:
    """
    wb1 = xlsxwriter.Workbook(filename)
    # 创建一个sheet工作对象
    ws = wb1.add_worksheet()
    for a in range(len(date_src)):
        for b in range(len(date_src[a])):
            c = date_src[a][b]
            ws.write(a, b, c)
    wb1.close()


if __name__ == '__main__':
    x = read_info("all.xlsx")
    date_list = get_useless(x)
    # date_list = del_useless(x)
    to_new_excel('final.xlsx', date_list)
