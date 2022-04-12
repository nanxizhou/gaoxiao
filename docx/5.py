# 要使用的知识点有：
# 1）添加 Paragraph 对象：Document 对象.add_paragraph()
# 2）添加 Run 对象：Paragraph 对象.add_run()
# 3）添加文字：Run 对象.add_text()
# 4）添加图片：Run 对象.add_picture()
from docx import Document

# 创建空的 Document 对象
doc = Document()

# 添加"两个黄鹂鸣翠柳，一行白鹭上青天。"
# 添加第一个 Paragraph 对象
para_1 = doc.add_paragraph()
# 添加 Run 对象
run_1 = para_1.add_run()
# 给 Run 对象添加文字
run_1.add_text('两个黄鹂鸣翠柳，一行白鹭上青天。')

# 添加"窗含西岭千秋雪，门泊东吴万里船。"
# 添加第二个 Paragraph 对象
para_2 = doc.add_paragraph()
# 添加 Run 对象
run_2 = para_2.add_run()
# 给 Run 对象添加文字
run_2.add_text('窗含西岭千秋雪，门泊东吴万里船。')
# 添加"唐-杜甫"和图片
# 添加第三个 Paragraph 对象
para_3 = doc.add_paragraph()
# 添加 Run 对象
run_3 = para_3.add_run()
# 给 Run 对象添加文字
run_3.add_text('唐-杜甫')
# 给 Run 对象添加图片，图片路径为'./杜甫.png'
run_3.add_picture('../docx/杜甫.png')

# 保存文档
doc.save('./绝句-杜甫2.docx')