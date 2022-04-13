# 读取工作表文件数据
with open('./04_月考勤表.xlsx', 'rb') as f:
    file_data = f.read()

# 设置内容类型为附件
attachment = MIMEText(file_data, 'base64', 'utf-8')

# 设置附件标题以及文件类型
attachment.add_header('Content-Disposition', 'attachment', filename='04_月考勤表.xlsx')