from Calculation.计算海峡报销和消费单 import *
from Calculation.计算合同制报销和消费单 import *
hx = haixia()
ht = hetong()


# 计算海峡的报销明细单和消费明细单
def all_procedures(filename1, filename2, filename3, filename4, filename5,
         sheet1, sheet2, sheet3):
    hx.haixia_reimbur_list(filename1, filename2, filename3, filename4,
                        sheet1, sheet2, sheet3)
    hx.haixia_deatcons_list(filename1, filename2, filename3, filename5,
                         sheet1, sheet2, sheet3)


all_procedures(filename1='3月员工资金变动流水（出账）.xls', filename2='3月林倩名单.xlsx',filename3='3月张怡君名单.xlsx',
               filename4='3月海峡报销明细单.xlsx', filename5='3月海峡消费明细单.xlsx',
               sheet1='员工资金变动流水', sheet2='Sheet1', sheet3='Sheet2'
)

# 完整计算程序，计算海峡和合同制的消费明细单和报销明细单
# def all_procedures(filename1, filename2, filename3, filename4, filename5, filename6, filename7,
#          sheet1, sheet2, sheet3):
#     hx.haixia_reimbur_list(filename1, filename2, filename3, filename4,
#                         sheet1, sheet2, sheet3)
#     hx.haixia_deatcons_list(filename1, filename2, filename3, filename5,
#                          sheet1, sheet2, sheet3)
#     ht.hetong_reimbur_list(filename1, filename2, filename3, filename6,
#                         sheet1, sheet2, sheet3)
#     ht.hetong_deatcons_list(filename1, filename2, filename3, filename7,
#                         sheet1, sheet2, sheet3)
#
#
# all_procedures(filename1='2月份员工资金变动流水（出账）.xls', filename2='2月林倩名单.xlsx',filename3='2月张怡君名单.xlsx',
#      filename4='海峡报销明细单.xlsx', filename5='海峡消费明细单.xlsx',
#     filename6='合同制报销明细单.xlsx', filename7='合同制消费明细单.xlsx',
#      sheet1='员工资金变动流水', sheet2='Sheet1', sheet3='Sheet2'
# )

