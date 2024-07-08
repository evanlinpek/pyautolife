import pandas as pd
import logging as log

df = pd.read_excel("team.xlsx")  # 读取excel文档，底层依赖openpyxl
# 查看数据
# print(df)  # 查看表格整体情况
# print(df.head())  # 查看前部5条
# print(df.tail())  # 查看尾部5条
# print(df.sample(5))  # 随机查看5条

# # 验证数据
# print(df.shape)  # 查看行数和列数 (100,6)
# df.info()  # 查看索引、数据类型和内存信息
# print(df.describe())  # 查看数值型列的汇总统计
# print(df.dtypes)  # 查看各字段类型
# print(df.axes)  # 显示数据行和列名
# print(df.columns)

# 建立索引
df.set_index('name',inplace=True)
# print(df.head())

# 查看指定列
# print(df['Q1'])
# 选择多列
# print(df[['team','Q1']])
# print(df.loc[:,['team','Q1']])  # df.loc[x,y],x 代表行，y，代表列。

# 选择行
# print(df[df.index == 'Liver'])  # 指定姓名
# print(df[0:3])  # 取前三行
# print(df[0:10:2])  # 在前10个中每两个取一个
# print(df.iloc[:10, :])  # 取前10行数据

# 指定行和列，同时给定行和列的显示范围
print(df.loc['Ben','Q1':'Q4'])  # 只看
print()