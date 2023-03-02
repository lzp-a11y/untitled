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

def get_result(value2, num_x_list):
    result = []
    for i in range(1, len(num_x_list)):
        iter = itertools.combinations(num_x_list, i)    # 对列表num_x_list里面的数据进行组合，看看有几种组合部分
        group_item = list(iter)                 # 将组合结果数值变成列表
        for j in range(1, len(group_item)):
    #         print(group_item[x], "之和等于", sum(group_item[x]))
    #         if sum(group_item[x]) == value2:
    #             result.append(group_item[x])
    # print('有', len(result), '种加法组合,得出', value2, '的组合是:')
            if sum(group_item[j]) in range(value2-20, value2):
                result.append(group_item[j])
    print('有', len(result), '种加法组合,接近', value2, '的组合是:')
    for i in range(len(result)):
        for j in range(len(result[i])):
            print(result[i][j], end=' ')
        print("和为", sum(result[i]))


if __name__ == '__main__':
    get_result(value2, num_x_list)



