import os
import time
import re
import pandas as pd
start_time = time.time()
root_dir = os.path.dirname(__file__)                 # 获取当前项目所在目录地址
excl1 = os.path.join(root_dir, "1月员工资金流水.xls")   # 拼接excel表格地址
excl2 = os.path.join(root_dir, "林倩名单.xlsx")
excl3 = os.path.join(root_dir, "张怡君名单.xlsx")


df1 = pd.read_excel(excl1, sheet_name='员工资金变动流水', skiprows=1)   # 省略指定行数，省略第一行
df2 = pd.read_excel(excl2, sheet_name='Sheet1')
df3 = pd.read_excel(excl3, sheet_name='Sheet1')

# how 合并方式，left左连接，保留left的全部数据， on： 链接的列属性。
result = pd.merge(df1, df2.loc[:, ['姓名', '人员类型', '现部门']], how='left', on='姓名')
result2 = pd.merge(result, df3.loc[:, ['姓名', '身份证', '专业线', '用工性质']], how='left', on='姓名')
filter1 = (result2["人员类型"] == '紧密型')          # 过滤人员类型列 等于紧密型的数据
filter2 = (result2["人员类型"] == '客服外包')
filter3 = (result2["人员类型"] == '营业外包')
filter4 = (result2["人员类型"] == '政企外包')
filter5 = (result2["状态"] == '出账')
result3 = result2.loc[filter1 | filter2 | filter3 | filter4]   # 把筛选后的数据赋值给result3
result3 = result3.loc[filter5]                              # 在result3的基础上进行筛选出账
result3 = pd.DataFrame(result3)                             # 将result3保存成数据框架后才能保存到excel

date = result3.iloc[0, 10]
result3['日期2'] = ''              # 18
result3['时间'] = ''               # 19
result3['餐类'] = ''               # 20
result3['补贴'] = ''               # 21
result3['员工消费金额'] = ''        # 22

for i in range(0, len(result3.index)):
    date = result3.iloc[i, 10]
    pattern1 = r"(\d{4}-\d{1,2}-\d{1,2})"
    pattern1 = re.compile(pattern1)
    pattern2 = r"(\d{1,2}:\d{1,2}:\d{1,2})"
    pattern2 = re.compile(pattern2)
    str_date1 = pattern1.findall(date)            # 获取分离后的日期2023-01-31
    str_date1 = str_date1[0]
    str_date2 = pattern2.findall(date)            # 获取分离后的时间18:18:13
    str_date2 = str_date2[0]
    result3.iloc[i, 18] = str_date1
    result3.iloc[i, 19] = str_date2

# 根据时间段区分早中晚餐
for i in range(0, len(result3.index)):
    date = result3.iloc[i, 19]
    if '07:00:00' < date < '09:00:00':
        result3.iloc[i, 21] = 2
        result3.iloc[i, 20] = '早餐'
    elif '10:00:00' < date < '13:30:00':
        result3.iloc[i, 21] = 6
        result3.iloc[i, 20] = '午餐'
    elif '17:00:00' < date < '19:30:00':
        result3.iloc[i, 21] = 2
        result3.iloc[i, 20] = '晚餐'

result4 = result3.drop_duplicates(['姓名', '日期2', '餐类'])   # 根据这3个字段去重

result5 = result4.drop_duplicates(['姓名'])                  # 根据名字去重，得到完整的名单
result5 = result5[['姓名']+['身份证']+['用工性质']+['现部门']+['专业线']+['员工消费金额']]   # 提取字段到新表格

# result4是去重后的表格
# result5是根据名字去重后的完整名单表格
for i in range(0, len(result5.index)):
    a = []
    name2 = result5.iloc[i, 0]
    for j in range(0, len(result4.index)):
        name1 = result4.iloc[j, 2]
        if name1 == name2:
            qian = result4.iloc[j, 21]
            a.append(qian)
    result5.iloc[i, 5] = sum(a)

result6 = result5[(result5['专业线'] != "产互")]           # result6专业线不等于产互的数据
result6 = result6.copy()                                 # 复制result6的数据，在复制的数据上进行操作
sum1 = result6['员工消费金额'].sum()                       # 对员工消费金额这列求和
result6.loc[len(result6)] = ['合计', '', '', '', '', sum1]
result7 = result5[(result5['专业线'] == "产互")]           # result7专业线等于产互的数据
result7 = result7.copy()                                 # 复制result6的数据，在复制的数据上进行操作
sum2 = result7['员工消费金额'].sum()                       # 对员工消费金额这列求和
result7.loc[len(result7)] = ['合计', '', '', '', '', sum2]

# 保存数据到同一个excel的不同不页中
excl4 = os.path.join(root_dir, "报销明细单.xlsx")
new_wb = pd.ExcelWriter(excl4)                    # 使用ExcelWriter()可以向同一个excel的不同sheet中写入对应的表格数据
result6.to_excel(new_wb, sheet_name='非合同制主业', index=False)     # 保存result6的数据到非合同制主业 页
result7.to_excel(new_wb, sheet_name='非合同制产互', index=False)     # 保存result7的数据到非合同制产互 页
new_wb.close()            # 直接调用关闭接口就可以了，close方法里面有save保存函数
end_time = time.time()
run_time = end_time - start_time


print("计算结束！")
print("程序运行时间：", int(run_time), "秒")

