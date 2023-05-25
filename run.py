import utils
import os
from read_excel import HandleExcel

print("文件格式只支持xlsx，请做好文件转化")
dir = input("请输入excel所在目录: ")
save_file_name = input("请输入需要整合到哪个文件: ")

if '.xlsx' not in save_file_name:
    print("文件格式只支持xlsx，请做好文件转化")
    os._exit()

export_column = input("请输入其他文件需要导出哪个列: ")

save_file = HandleExcel(dir + save_file_name)

for file in utils.findAllFile(dir):
    if not os.path.isfile or save_file_name in file or '.xlsx' not in file:
        continue
    
    read_file = HandleExcel(dir + file)
    
    data = read_file.read_all_col(export_column)

    save_file.write_list_to_column(data)

print("导出完成")
