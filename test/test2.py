from openpyxl import load_workbook

# 打开工作表
file_path = './material/事业01部_副本.xlsx'
wb = load_workbook(file_path)
ws = wb.active

# 调整第二列列宽
ws.column_dimensions['B'].width = 20

# 保存
wb.save(file_path)
