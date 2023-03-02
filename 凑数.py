import itertools
import pandas as pd
import numpy as np
import xlrd
import openpyxl
workbook = openpyxl.load_workbook(r"E:\Desktop\工作簿1.xlsx")
sheet = workbook['Sheet1']
max_row = sheet.max_row
num_x_list = []
for i in range(2, max_row+1):
    value1 = sheet["A{}".format(i)].value  # 取出第一列的数值
    if value1 != 0:
        num_x_list.append(value1)
value2 = sheet["B2"].value
value3 = sheet["B3"].value


result = []
for i in range(1, len(num_x_list)):
    iter = itertools.combinations(num_x_list, i)    # 对列表num_x_list里面的数据进行组合，看看有几种组合部分
    group_item = list(iter)                         # 将所有组合结果数值变成列表
    for j in range(1, len(group_item)):             # 遍历所有可能性的组合，然后求和
#         print(group_item[j], "之和等于", sum(group_item[j]))
#         if sum(group_item[j]) == value2:
#             result.append(group_item[x])
# print('有', len(result), '种加法组合,得出', value2, '的组合是:')
        if sum(group_item[j]) in range(value2-5, value2+1):
            result.append(group_item[j])
print('有', len(result), '种加法组合,接近', value2, '的组合是:')
for i in range(len(result)):
    print(result[i], '和为：', sum(result[i]))
    # 计算组合内数值之和
    list1 = list(result[i])            # list1为符合第一个数值的数字组合。
    list2 = list(set(num_x_list) - set(list1))    # list2为扣掉符合第一个数值组合后的数字，因为一个数字只能用一次
    result2 = []
    for i2 in range(1, len(list2)):
        iter2 = itertools.combinations(list2, i2)         # 对列表list2里面的数据进行组合，看看有几种组合部分
        group_item2 = list(iter2)                        # 将组合结果数值变成列表
        for j2 in range(1, len(group_item2)):             # 遍历所有可能性的组合，然后求和
            if sum(group_item2[j2]) in range(value3-5, value3+1):
                result2.append(group_item2[j2])
    print('有', len(result2), '种加法组合,接近', value3, '的组合是:')
    for i in range(len(result2)):
        print(result2[i], '和为：', sum(result2[i]))


