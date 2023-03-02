import pandas as pd
import os
import time
start_time = time.time()
root_dir = os.path.dirname(__file__)                 # 获取当前项目所在目录地址
excl1 = os.path.join(root_dir, "1月员工资金流水.xls")   # 拼接excel表格地址
excl2 = os.path.join(root_dir, "林倩名单.xlsx")
excl3 = os.path.join(root_dir, "张怡君名单.xlsx")
excl4 = os.path.join(root_dir, "合同制员工刷卡次数统计.xlsx")

df1 = pd.read_excel(excl1, sheet_name='员工资金变动流水', skiprows=1)   # 省略指定行数，省略第一行
df2 = pd.read_excel(excl2, sheet_name='Sheet1')               # 林倩名单
df3 = pd.read_excel(excl3, sheet_name='Sheet1')               # 张怡君名单

# how 合并方式，left左连接，保留left的全部数据， on： 链接的列属性。
result = pd.merge(df1, df2.loc[:, ['姓名', '人员类型', '现部门', '二级机构']], how='left', on='姓名')
result2 = pd.merge(result, df3.loc[:, ['姓名', '身份证', '专业线', '用工性质']], how='left', on='姓名')
filter1 = (result2["人员类型"] == '紧密型')          # 过滤人员类型列 等于紧密型的数据
filter2 = (result2["人员类型"] == '客服外包')
filter3 = (result2["人员类型"] == '营业外包')
filter4 = (result2["人员类型"] == '政企外包')
filter5 = (result2["人员类型"] == '合同制')
filter6 = (result2["现部门"] == '工业互联网BU')
filter7 = (result2["现部门"] == '云网中心')
filter8 = (result2["现部门"] == '云网中心（借调）')
filter9 = (result2["二级机构"] == '交付组')
filter10 = (result2["状态"] == '出账')

result3 = result2.loc[filter5 & filter6 & filter10]   # 人员类型是合同制，现部门是工业互联网BU，状态出账
# 人员类型是合同制，现部门是云网中心或者云网中心（借调），并且二级机构为交付组且状态为出账
result4 = result2.loc[(filter7 | filter8) & filter5 & filter9 & filter10]
result5 = pd.concat([result3, result4], ignore_index=True)   # 拼接result3和4，组成5,5位产互的数据
result6 = result2.loc[filter5 & filter10]      # 合同制且状态为出账
result7 = pd.concat([result6, result5, result5]).drop_duplicates(keep=False)

new_wb = pd.ExcelWriter(excl4)                    # 使用ExcelWriter()可以向同一个excel的不同sheet中写入对应的表格数据
result5.to_excel(new_wb, sheet_name='产互刷卡次数', index=False)     # 保存result5的数据到非合同制主业 页
result7.to_excel(new_wb, sheet_name='非产互刷卡次数', index=False)     # 保存result6的数据到非合同制主业 页
new_wb.close()            # 直接调用关闭接口就可以了，close方法里面有save保存函数
end_time = time.time()
run_time = end_time - start_time

print("合同制产互刷卡次数为：{}".format(len(result5)))
print("合同制非产互刷卡次数为：{}".format(len(result7)))
print("合同制员工刷卡总次数为：{}".format(len(result5)+len(result7)))
print("计算结束！")
print("程序运行时间：", int(run_time), "秒")

