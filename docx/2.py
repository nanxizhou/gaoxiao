from docx import Document

doc = Document('../工作/涨薪通告-练习/康志威涨薪通告.docx')

# 循环遍历 Document 对象中的每一个 Paragraph 对象
for para in doc.paragraphs:
    # 打印 Paragraph 对象中的文字
    print(para.text)

# 打印 Paragraph 对象的个数
print(len(doc.paragraphs))