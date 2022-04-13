from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference

# 读取工作簿
wb = load_workbook('../material/事业01部_副本.xlsx')
# 读取工作簿中的活跃工作表
ws = wb.active
# 实例化 LineChart() 类，得到 LineChart 对象
chart = LineChart()
# 引用工作表的部分数据
data = Reference(worksheet=ws, min_row=3, max_row=5, min_col=2, max_col=3)
# 添加被引用的数据到 LineChart 对象，设置参数from_rows的参数为False
chart.add_data(data, from_rows=True, titles_from_data=False)
# 添加 LineChart 对象到工作表中，指定生成折线图的位置
ws.add_chart(chart, "C12")

# 保存文件
wb.save('../material/事业03部_副本.xlsx')
# 打印'代码运行成功!'
print('代码运行成功!')

