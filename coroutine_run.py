import asyncio
import time
import utils
import os
import utils
from read_csv import HandleCsv
from read_excel import HandleExcel

dir = "C:\\Users\\Administrator\\Desktop\\曾宇翔资料\\数据处理-zyx\\20230522\\sample6\\16-4"
export_column = "D"
save_file_name = "total.xlsx"
write_col = 1

def get_all_file():
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
    return {**all_num_files, **other_files}

async def get_data(index, filename):
    read_file = HandleCsv(filename)
    
    data = read_file.read_all_col(export_column)

    return index, data

async def main():
    start_time = time.time()

    all_files = get_all_file()
    task = [get_data(i, file) for i, file in all_files.items()]
    rel = await asyncio.gather(*task)
    global write_col

    save_file = HandleExcel(dir  + "\\" + save_file_name)
    for data in rel:
        i, d = data
        if isinstance(i, int):
            d.insert(0, str(i) + "°")
        else:
            d.insert(0, i)

        save_file.write_list_to_column(d, write_col)
        write_col += 1

    print("Execution Time:", time.time() - start_time)

asyncio.run(main())
