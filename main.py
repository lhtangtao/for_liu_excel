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
import os

if __name__ == '__main__':
    allxls = []
    base = os.path.dirname(__file__)
    all_excels = os.path.join(base, 'excels')
    for root, dirs, files in os.walk(all_excels):
        allxls = files  # 当前路径下所有非目录子文件
    merge(allxls, 'all.xlsx')  # 合并列表中的excel到一个叫做all.xlsx的excel文件夹中。
    total_base_dir = (os.path.dirname(__file__) + '/all.xlsx').replace('/', "\\")
    sort_by('C1', total_base_dir)  # 根据C行来进行排序
    x = read_info("all.xlsx")  # 从all.xlsx中读取数据
    date_list = get_useless(x)  # 把所有的重复使用的东西写入到一个列表中
    to_new_excel('final.xlsx', date_list)  # 把列表中的信息写入到新的excel中