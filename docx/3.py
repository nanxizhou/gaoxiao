import os
# 导入 Document 类
from docx import Document

# 设置目标文件夹路径
path = "../工作/涨薪通告-练习/"

# 获取目标文件夹下的所有文件名
file_list = os.listdir(path)

for file in file_list:
    # 拼接文件路径
    file_path = path + file
    # 打开 Word 文件
    doc = Document(file_path)
    # 打印文档对象
    print(doc)