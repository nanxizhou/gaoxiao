from openpyxl import load_workbook

# 打开工作簿
wb = load_workbook("./04_月考勤表.xlsx")
# 打开工作表
ws = wb.active
# 打印表格
print(ws)