import time
import utils
import os
from read_csv import HandleCsv
from read_excel import HandleExcel

print("数据文件仅支持csv格式，请做好格式转化")
dir = input("请输入数据文件所在目录: ")
dir = dir.strip()
export_column = input("请输入被整合文件需要导出哪个列: ")
save_file_name = input("请输入需要整合到哪个文件: ")
save_file_name = save_file_name.strip()
write_col = input("请输入从整合文件的哪一列开始导入: ")
write_col = utils.col_to_number(write_col) + 1

# dir = "C:\\Users\\Administrator\\Desktop\\曾宇翔资料\\数据处理-zyx\\20230522\\sample6\\16-4"
# export_column = "D"
# save_file_name = "total.xlsx"
# write_col = 1

if '.xlsx' not in save_file_name:
    print("整合文件格式只支持xlsx，请做好文件转化")
    os._exit(1)

save_file = HandleExcel(dir  + "\\" + save_file_name)

all_num_files = {}
other_files = {}
for file in utils.findAllFile(dir):
    filename, ext = os.path.splitext(os.path.basename(file))
    
    if not os.path.isfile(file) or save_file_name in file or not ext == ".csv":
        continue

    if str.isdigit(filename):
        all_num_files[int(filename)] = file
    else:
        other_files[filename] = file

all_num_files = dict(sorted(all_num_files.items()))

all_files = {**all_num_files, **other_files}

# start_time = time.time()
for i, file in all_files.items():
    read_file = HandleCsv(file)
    
    data = read_file.read_all_col(export_column)

    if isinstance(i, int):
        data.insert(0, str(i) + "°")
    else:
        data.insert(0, i)

    save_file.write_list_to_column(data, write_col)
    write_col += 1
# print("Execution Time:", time.time() - start_time)
print("导出完成")
