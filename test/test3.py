# 导入模块
import os
# 设置文件夹路径
path = './各部门利润表汇总/'
# 返回当前目录下所有文件名
files = os.listdir(path)
# 循环文件名列表
for file in files:
    # 拼接文件路径
    file_path = path + file
    # 打印文件路径
    print(file_path)
