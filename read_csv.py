import csv
import utils

class HandleCsv():

    def __init__(self, filename):
        self.filename = filename

    def read_all_col(self, column_name):
        col_num = utils.col_to_number(column_name)
        data = []
        with open(self.filename, 'r') as file:
            reader = csv.reader(file)
            next(reader)

            for row in reader:
                data.append(eval(row[col_num]))

        return data

if __name__ == '__main__':
    file = input("请输入文件: ")
    handle = HandleCsv(file.strip())
    print(handle.read_all_col("D"))
