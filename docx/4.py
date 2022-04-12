from docx import Document

doc = Document('../工作/涨薪通告-练习/康志威涨薪通告.docx')

# 循环遍历 Document 对象中的每一个 Paragraph 对象
for para in doc.paragraphs:
    # 循环遍历 Paragraph 对象中的每一个 Run 对象
    for run in para.runs:
        # 打印 Run 对象
        print(run)
        # 打印 Run 对象中的文字
        print(run.text)