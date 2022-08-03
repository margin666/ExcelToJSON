import os
import json
import openpyxl

class Compiler(object):
    """转换类"""
    def __init__(self, fileName, resultName) -> None:
        self.filePath = self.resolvePath(fileName)
        self.resultPath = self.dist(resultName)
    # 获取文件
    def resolvePath(self,name):
        return os.path.join(os.getcwd(), 'excels', name)
    # 导出json文件
    def dist(self, resultName):
        return os.path.join(os.getcwd(), 'dist', resultName)
    # 获取sheet的第一行数据   return dict
    def dataTitleDict(self, sheet):
        data = {}
        for cell in sheet[1]:
            data[cell.value] = ''
        return data

    # 根据就获取的dict，组装list并return
    def dataContentList(self, dataTitle, sheet):
        m = []
        listKeys = list(dataTitle.keys())
        for row in sheet.iter_rows(min_row=2):
            temp = dict(dataTitle)
            for i,cell in enumerate(row):
                key = listKeys[i]
                temp[key] = cell.value
            m.append(temp)
        return m
    def run(self):
        worker_book = openpyxl.load_workbook(self.filePath)
        data = {}
        for sheet in worker_book:
            keys = self.dataTitleDict(sheet)
            list = self.dataContentList(keys, worker_book.active)
            data[sheet.title] = list
        f = open(self.resultPath,'w',encoding='utf-8')
        f.write(json.dumps(data))
        f.close()





