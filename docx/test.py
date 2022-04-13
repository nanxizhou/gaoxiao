# 导入库和模块
import os
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt

# 设置目标文件夹路径为'../工作/涨薪通告-实操/'
path = '../工作/涨薪通告-实操/'

# 获取目标文件夹下的所有文件名
file = os.listdir(path)
# 循环遍历所有文件名
for row in file:
    # 拼接文件路径
    file_path = path + row
    # 打开 Word 文件
    doc = Document(file_path)

    # 添加文字'盖章：'和公司电子章图片(图片路径为：'./Shining.png')
    para_1 = doc.add_paragraph('盖章')
    # 添加Run对象 run_stamp
    run_stamp = para_1.add_run()
    # 图片
    run_stamp.add_picture('./Shining.png')
    # 添加公司名称'闪光科技金融公司(Shining Fintech Company)'并将字号设置为四号，文字靠右对齐并加粗
    para_2 = doc.add_paragraph()
    #添加Run对象
    run_comp = para_2.add_run("闪光金融公司(Shining Finance Company)")
    # 对齐
    para_2.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    # 四号字体
    run_comp.font.size = Pt(4)
    # 加粗
    run_comp.font.bold = True
    # 保存 Word 文件
    doc.save(file_path)