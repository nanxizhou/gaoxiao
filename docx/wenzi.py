# 添加文字需要 Run 对象调用方法add_text(text)
from docx import Document
doc = Document()
para_1 = doc.add_paragraph()
run_1 = para_1.add_run()
run_1.add_text('使用run添加文字')
doc.save('./添加文字.docx')
