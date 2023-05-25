from openpyxl import *

class HandleExcel():
    def __init__(self, filenname):
        self.filename = filenname
        self.handle = load_workbook(filenname)
        self.active_sheet = self.handle.active

    # Read all values in a column    
    def read_all_col(self, column_name):
        max_row = self.active_sheet.max_row
        column_data = []

        for i in range(1, max_row + 1):
            column_data.append(self.active_sheet[column_name + str(i)].value)

        return column_data
    
    def write_list_to_column(self, data):
        max_col = self.active_sheet.max_column + 1
        print(max_col)

        for i, item in enumerate(data):
            self.active_sheet.cell(column=max_col, row=i+1).value = item

        self.handle.save(self.filename)

if __name__ == '__main__':
    read_file = HandleExcel('/Users/zhaoji/Downloads/b.xlsx')
    data = read_file.read_all_col('D')

    write_file = HandleExcel('/Users/zhaoji/Downloads/a.xlsx')
    write_file.write_list_to_column(data)
