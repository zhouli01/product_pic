# 库导入
import os
from matplotlib import pyplot as plt
import numpy as np
# 卸载新版本 pip uninstall xlrd
# 安装老版本：pip install xlrd=1.2.0 (或者更早版本)

import xlrd
import os
from datetime import datetime
from xlrd import xldate_as_tuple

excel_path = os.path.join(os.getcwd(), 'data.xlsx')
print('Excel文件的路径：' + excel_path)
excel_file = xlrd.open_workbook(excel_path)
table = excel_file.sheets()[0]

print('已经打开的工作簿的名字：' + table.name)
print('**********开始读取Excel单元格的内容**********')

all_content = []
for i in range(table.nrows):
    row_content = []
    for j in range(table.ncols):
        ctype = table.cell(i, j).ctype  # 获取单元格返回的数据类型
        cell_value = table.cell(i, j).value  # 获取单元格内容
        if ctype == 2 and cell_value % 1 == 0:  # 是否是数字类型
            cell_value = int(cell_value)
        elif ctype == 3:  # 是否是日期
            date = datetime(*xldate_as_tuple(cell_value, 0))
            cell_value = date.strftime('%Y/%m/%d %H:%M:%S')
        elif ctype == 4:  # 是否是布尔类型
            cell_value = True if cell_value == 1 else False
        row_content.append(cell_value)
    all_content.append(row_content)
    # print('[' + ', '.join("'" + str(element) + "'" for element in row_content) + ']')
    print("row_content:", row_content)
    # print(all_content)
print('**********Excel单元格的内容读取完毕**********')

print('行数:%d' % table.nrows)  # 打印行数
print('列数:%d' % table.ncols)  # 打印列数


print('========================')
# print('第二行的内容：' + str(table.row_values(1)))  # 打印一行的内容
# print('第二列的内容：' + str(table.col_values(1)))  # 打印一列的内容

countries =[]
gold_medal =[]
silver_medal=[]
bronze_medal = []

nrows = table.nrows
ncols = table.ncols

# 国家和奖牌数据读取
for i in range(1, nrows - 1):
    for j in range(0, ncols):
        if j == 0:
            countries_enu = table.cell(i, 0).value
            countries.append(str(countries_enu))
        if j == 1:
            gold_medal_enu = table.cell(i, 1).value
            gold_medal.append(int(gold_medal_enu))
        if j == 2:
            silver_medal_enu = table.cell(i, 2).value
            silver_medal.append(int(silver_medal_enu))
        if j == 3:
            bronze_medal_enu = table.cell(i, 3).value
            bronze_medal.append(int(bronze_medal_enu))

# print("countries:",countries)
# print("gold_medal:",gold_medal)
# print("silver_medal:",silver_medal)
# print("bronze_medal:",bronze_medal)



# 参数设置
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 120
plt.rcParams['figure.figsize'] = (5, 3)

# 将横坐标国家转换为数值
x = np.arange(len(countries))
width = 0.2

# 计算每一块的起始坐标
gold_x = x
silver_x = x + width
bronze_x = x + 2 * width

# 绘图
# 柱状图
plt.bar(gold_x, gold_medal, width=width, color="gold", label="王娇")
plt.bar(silver_x, silver_medal, width=width, color="silver", label="张锐娟")
plt.bar(bronze_x, bronze_medal, width=width, color="saddlebrown", label="樊韦秀")

# 折线图
# plt.plot(silver_medal,color="silver",label="张锐娟")
# plt.plot(bronze_medal,color="saddlebrown",label="樊韦秀")

# 将横坐标数值转换为国家
plt.xticks(x + width, labels=countries)

# 显示柱状图的高度文本
for i in range(len(countries)):
    plt.text(gold_x[i], gold_medal[i], gold_medal[i], va="bottom", ha="center", fontsize=8, color="gold")
    plt.text(silver_x[i], silver_medal[i], silver_medal[i], va="bottom", ha="center", fontsize=8, color="silver")
    plt.text(bronze_x[i], bronze_medal[i], bronze_medal[i], va="bottom", ha="center", fontsize=8, color="saddlebrown")

plt.title("催费统计信息", x=0.5, y=1.1)
# plt.subplots_adjust(left=0.1, right=0.6)
# 显示图例
plt.legend(loc="upper right", bbox_to_anchor=(1, 1.1), ncol=3)

#保存图片
plt.savefig('./Picture/20240531.png', bbox_inches="tight", dpi=120)

#图片展示预览
plt.show()
