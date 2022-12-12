
import xlwings as xw
import os
import glob
import sqlite3
from functools import wraps
from openpyxl.drawing.image import Image

HDB = None

def DBDecorator(func):
    """
    하위 func 실행여부도 기록됨
    """
    @wraps(func)
    def wrapper(self, *args):
        global HDB
        HDB = DBManager()
        ret = func(self, *args)
        HDB.close()
        return ret
    return wrapper

class DBManager:
    def __init__(self):
        if os.path.isdir('DataBase') != True:
            os.makedirs('DataBase')
        self.dbpath = "D:/ssssssssssss/DataBase/ExcelRPA.db"
        self.dbConn = sqlite3.connect(self.dbpath, isolation_level = None)
        self.c = self.dbConn.cursor()

    def close(self):
        self.dbConn.close()

@DBDecorator
def db_select(cmd: str):
    HDB.c.execute(cmd)
    result = HDB.c.fetchall()
    return result

path = r"C:\Users\9350816\Desktop\다국어자동화(4).xlsx"
 
wb = xw.Book(path)
print(wb.name)

# 모든 시트명 표시
sheets = wb.sheets
li_sheets = [s.name for s in sheets]
print(li_sheets)

langList = ["중국어"]
img_path = os.path.join(os.getcwd(), 'logo.png')

for idx, lang in enumerate(langList):
    if lang in li_sheets:
        wb.sheets(lang).delete()
    ws = wb.sheets.add(lang)

    dataList = db_select(f"SELECT * FROM '{lang}'")
    for i, data in enumerate(dataList):
        if i == 0:
            # try:
            # img = Image(data[0])
            # img.width = 400
            # img.height = 155
            ws.pictures.add(img_path)
            # except:
            #     ws.range(f'A{i+1}')
        ws.range(f'A{i+2}').value=data
        print(list(data))

wb.save(r"C:\Users\9350816\Desktop\다국어자동화(4).xlsx")
wb.close()