import itertools
import pandas as pd
import numpy as np
import xlrd
import openpyxl
import os
from datetime import datetime


def coushu(filename1, filename2, data1, data2, sheet_name1, to_sheet_name1, to_sheet_name2):
    root_dir = os.path.dirname(__file__)                 # 获取当前项目所在目录地址
    excl1 = os.path.join(root_dir, "{}.xlsx".format(filename1))        # 拼接excel表格地址
    excl2 = os.path.join(root_dir, "{}.xlsx".format(filename2))
    df1 = pd.read_excel(excl1, sheet_name='{}'.format(sheet_name1), skiprows=1)   # 省略指定行数，省略第一行, parse_dates=['时间']
    df1['凑值'] = None
    df1.iloc[0, 5] = '{}'.format(data1)
    df1.iloc[1, 5] = '{}'.format(data2)
    num_x_list = []
    for i in range(1, len(df1)):
        value1 = df1.iloc[i, 4]  # 取出第一列的数值
        value1 = int(value1)
        if value1 != 0:
            num_x_list.append(value1)
    value2 = int(df1.iloc[0, 5])         # 取出要凑的2个值
    value3 = int(df1.iloc[1, 5])

    result = []
    result2 = []
    for i in range(0, len(num_x_list)):
        iter = itertools.combinations(num_x_list, i)    # 对列表num_x_list里面的数据进行组合，看看有几种组合部分
        group_item = list(iter)                         # 将所有组合结果数值变成列表
        for j in range(0, len(group_item)):             # 遍历所有可能性的组合，然后求和
            if sum(group_item[j]) in range(value2-20, value2+1):
                result.append(group_item[j])
    print('有', len(result), '种加法组合,接近', value2, '的组合是:')
    for i in range(0, 1):
        print(result[i], '和为：', sum(result[i]))
        # 计算组合内数值之和
        list1 = list(result[i])            # list1为符合第一个数值的数字组合。
        list2 = list(set(num_x_list) - set(list1))    # list2为扣掉符合第一个数值组合后的数字，因为一个数字只能用一次
        for i2 in range(0, len(list2)):
            iter2 = itertools.combinations(list2, i2)         # 对列表list2里面的数据进行组合，看看有几种组合部分
            group_item2 = list(iter2)                        # 将组合结果数值变成列表
            for j2 in range(0, len(group_item2)):             # 遍历所有可能性的组合，然后求和
                if sum(group_item2[j2]) in range(value3-20, value3+1):
                    result2.append(group_item2[j2])
        print('有', len(result2), '种加法组合,接近', value3, '的组合是:')
        for i3 in range(len(result2)):
            print(result2[i3], '和为：', sum(result2[i3]))

    df2 = df1[df1["金额(元）"].isin(result[0])]    # 产互
    df2 = df2.copy()
    sum1 = df2['金额(元）'].sum()
    df2 = df2[['序号']+['时间']+['货名']+['单价']+['金额(元）']]
    df2.loc[300] = ['合计', '', '', '', sum1]

    df3 = df1[df1["金额(元）"].isin(result2[0])]   # 主页
    df3 = df3.copy()
    sum2 = df3['金额(元）'].sum()
    df3 = df3[['序号']+['时间']+['货名']+['单价']+['金额(元）']]
    df3.loc[300] = ['合计', '', '', '', sum2]

    new_wb = pd.ExcelWriter(excl2)                    # 使用ExcelWriter()可以向同一个excel的不同sheet中写入对应的表格数据
    df2.to_excel(new_wb, sheet_name='{}'.format(to_sheet_name1), index=False)     # 产互
    df3.to_excel(new_wb, sheet_name='{}'.format(to_sheet_name2), index=False)     # 主业
    new_wb.close()            # 直接调用关闭接口就可以了，close方法里面有save保存函数

"""
filename1: 取值的文件名    filename2: 要保存的结果文件     data1: 要凑值的数据1（产互） data2: 要凑值的数据2（主业）
sheet_name1：取值文件的sheet名     to_sheet_name1：保存结果文件的sheet名（产互）  to_sheet_name2：保存结果文件的sheet名（主业）
"""
coushu(filename1="2023年2月农品优农采购明细表7", filename2='肉类',
       data1=3292, data2=10000, sheet_name1='蔬菜、肉等日常采购', to_sheet_name1='肉类产互', to_sheet_name2='肉类主业')

