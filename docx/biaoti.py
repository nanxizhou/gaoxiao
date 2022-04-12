# 添加标题（Heading 对象）的方法：Document 对象.add_heading()
from docx import Document
docx = Document()
docx.add_heading('标题零',  level=0)
docx.add_heading('标题一',  level=1)

docx.save('./标题.docx')