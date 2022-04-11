from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# 设置邮箱账号
account = input('请输入邮箱账户：')
# 设置邮箱授权码

# 实例化smtp对象，设置邮箱服务器，端口
smtp = smtplib.SMTP_SSL('smtp.qq.com', 465)
# 登录qq邮箱
smtp.login(account,  'jlrnctylmtmzdejg')

content = '这是王振义的迟到表格'
# 添加正文，创建简单邮件对象
email_content = MIMEText(content, 'plain', 'utf-8')
with open('../material/事业01部_副本.xlsx', 'rb') as f:
    file_data = f.read()
attachment = MIMEText(file_data, 'base64', 'utf-8')
attachment.add_header('Content-Disposition', 'attachment', filename='../material/事业01部_副本.xlsx.xlsx')

msg = MIMEMultipart()
msg.attach(email_content)
msg.attach(attachment)
# 设置发送者信息
msg['From'] = '陈知枫'
# 设置接受者信息
msg['To'] = '热爱学习的你'
# 设置邮件标题
msg['Subject'] = '来自知枫的一封信'

# 发送邮件
smtp.sendmail(account, '2818813711@qq.com', msg.as_string())
# 关闭邮箱服务
smtp.quit()
