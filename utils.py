import os

def findAllFile(base):
    for root, ds, fs in os.walk(base):
        for f in fs:
            fullname = os.path.join(root, f)
            yield fullname

def col_to_number(col):
    ab = '_ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    col = col.upper()
    w = 0
    for _ in col:
        w *= 26
        w += ab.find(_)
    return w - 1
            
if __name__ == '__main__':
    dir = input("请输入: ")
    for file in findAllFile(dir):
        print(file)
