from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
# 创建空的 Document 对象
doc = Document()

# 添加 Paragraph 对象
para_1 = doc.add_paragraph('两个黄鹂鸣翠柳，一行白鹭上青天。')
# 设置对齐方式为左对齐
para_1.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

# 添加 Paragraph 对象
para_2 = doc.add_paragraph('窗含西岭千秋雪，门泊东吴万里船。')
# 设置对齐方式为居中对齐
para_2.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


# 添加 Paragraph 对象
para_3 = doc.add_paragraph('唐-杜甫-《绝句》')
# 设置对齐方式为右对齐
para_3.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

doc.save('./对齐方式.docx')