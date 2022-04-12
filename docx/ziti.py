from docx import Document
# 导入控制对齐方式所需
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
# 导入控制字体大小所需
from docx.shared import Pt

# 设置文件路径
file_path = '../docx/马邦德涨薪通告.docx'

# 打开文档
doc = Document(file_path)
# 添加段落2
para = doc.add_paragraph()
# 设置对齐方式
para.paragraph_format.alignment=WD_PARAGRAPH_ALIGNMENT.RIGHT
# 添加 run_comp
run_comp = para.add_run("闪光金融公司(Shining Finance Company)")
# 设置字体大小为 14pt
run_comp.font.size=Pt(14)
# 设置字体加粗
run_comp.font.bold=True
# 保存文件
doc.save('./添加带样式的文字.docx')