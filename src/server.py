import openpyxl
import os
import json



def resolvePath(fileName):
    return os.path.join(os.getcwd(), 'excels', fileName)
def dist():
    return os.path.join(os.getcwd(), 'dist', 'result.json')

def read_file(url):
    return openpyxl.load_workbook(url)


def dataTitleDict(sheet):
    data = {}
    for cell in sheet[1]:
        data[cell.value] = ''
    return data

def dataContentList(dataTitle, sheet):
    
    m = []
    listKeys = list(dataTitle.keys())
    for row in sheet.iter_rows(min_row=2):
        
        temp = dict(dataTitle)
        for i,cell in enumerate(row):
            key = listKeys[i]
            temp[key] = cell.value
        m.append(temp)
    return m
            



def change():
    filePath = resolvePath('test.xlsx')
    worker_book = read_file(filePath)
    data = dataTitleDict(worker_book.active)
    datac = dataContentList(data, worker_book.active)
    # print(json.dumps(datac))
    f = open(dist(),'w',encoding='utf-8')
    f.write(json.dumps(datac))
    f.close()

change()
# 遍历所有的sheet
# 处理单个sheet
# 拿到第一行的表头，创建一个基本对象
# 创建返回值
# 循环除第一行外所有数据
# 将结果返回

