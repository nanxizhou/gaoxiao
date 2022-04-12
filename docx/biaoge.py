# 添加表格（Table 对象）的方法：Document 对象.add_table()
from docx import Document
# 实例Document对象
docx = Document()
# 添加table对象
table = docx.add_table(rows=5, cols=5, style='Table Grid')
# 给表格中添加内容
i = 1
j = 1
for row in table.rows:
    for cell in row.cells:
        cell.text = f'第{i}行、第{j}列'
        j += 1
    i += 1
    j = 1
docx.save('./表格.docx')
