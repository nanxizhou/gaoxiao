import os

# 设置目标文件夹路径
path = "./工作/涨薪通告-练习/"
# 获取目标文件夹下的所有文件名
file_list = os.listdir(path)
# 循环取出文件名
for file in file_list:
    # 拼接文件路径
    file_path = file+path
    # 打印文件路径
    print(file_path)
