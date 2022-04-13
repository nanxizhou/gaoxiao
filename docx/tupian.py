# Run 对象去调用add_picture(path)
from docx import Document
doc = Document()
para_1 = doc.add_paragraph()
run_1 = para_1.add_run()
run_1.add_picture('../docx/Shining.png')
doc.save('./添加图片.docx')