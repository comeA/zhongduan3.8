'''
重要说明：

开发人员 ： 卢鹤斌
  #为了测试 openpyxl 版本问题。因原来的项目无法使用最新版本的 openpyxl 版本：3.1.3

'''


import openpyxl
import sys

print(f"Python 解释器路径：{sys.executable}")
print(f"openpyxl 模块路径：{openpyxl.__file__}")
print(f"openpyxl 版本：{openpyxl.__version__}")

try:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([1, 2, 3])  # 添加一行数据，确保有数据可删除
    ws.delete_rows(min_row=2, max_row=ws.max_row)  # 确保只有这一行
    wb.save("test.xlsx")
    print("测试成功！")
except Exception as e:
    print(f"发生错误：{e}")

