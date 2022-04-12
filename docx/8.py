# 可能要使用的知识点有：
# 1）实例化 Document 类：Document()
# 2）添加 Paragraph 对象：Document 对象.add_paragraph(text)
# 3）添加 Run 对象：Paragraph 对象.add_run(text)
# 4）添加文字：Run 对象.add_text()
# 5）添加图片：Run 对象.add_picture()
from docx import Document

# 设置文件路径
file_path = '../docx/马邦德涨薪通告.docx'

# 打开 Word 文件
doc = Document(file_path)

# 添加'盖章：'与电子章图片(图片路径为：'./Shining.png')
# 添加段落1
para_1 = doc.add_paragraph('盖章')
# 添加 run_stamp
run_stamp = para_1.add_run()
run_stamp.add_picture('../docx/Shining.png')


# 保存文件
doc.save('./添加盖章和图片.docx')