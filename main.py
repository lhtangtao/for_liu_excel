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
from del_unuse import get_useless, read_info, to_new_excel
from merge_excel import merge
from sort_excel import sort_by

if __name__ == '__main__':
    allxls = ['1.xlsx', '2.xlsx', '3.xlsx', '4.xlsx', '5.xlsx', '6.xlsx', '7.xlsx', '8.xlsx', '9.xlsx',
              '10.xlsx', '11.xlsx', '101.xlsx', '102.xlsx', '103.xlsx', '104.xlsx', '105.xlsx', '106.xlsx',
              '107.xlsx']  # 此处输入你要合并的excel文件的文件名
    # allxls = ['55.xlsx', '66.xlsx']
    merge(allxls, 'all.xlsx')  # 合并列表中的excel到一个叫做all.xlsx的excel文件夹中。
    sort_by('C1', r'C:\Users\tangtao\PycharmProjects\for_liu_excel/all.xlsx')  # 根据C行来进行排序
    x = read_info("all.xlsx") # 从all.xlsx中读取数据
    date_list = get_useless(x) # 把所有的重复使用的东西写入到一个列表中
    to_new_excel('final.xlsx', date_list) # 把列表中的信息写入到新的excel中
