
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
        self.dbpath = "D:/new/DataBase/ExcelRPA.db"
        self.dbConn = sqlite3.connect(self.dbpath, isolation_level = None)
        self.c = self.dbConn.cursor()

    def close(self):
        self.dbConn.close()

@DBDecorator
def db_select(cmd: str):
    HDB.c.execute(cmd)
    result = HDB.c.fetchall()
    return result

@DBDecorator
def db_tables(cmd: str):
    HDB.c.execute(cmd)
    result = set([col_tuple[0] for col_tuple in HDB.c.description])
    return result

@DBDecorator
def db_columns(cmd: str):
    HDB.c.execute(cmd)
    result = [col_tuple[0] for col_tuple in HDB.c.description]
    return result

path = r"C:\Users\82109\Desktop\다국어자동화(4).xlsx"
 
wb = xw.Book(path)
print(wb.name)

# 모든 시트명 표시
sheets = wb.sheets
li_sheets = [s.name for s in sheets]
print(li_sheets)

langList = ["중국어"]


for idx, lang in enumerate(langList):
    CHR_COL = 65
    # 시트 있으면 삭제 후 생성
    if lang in li_sheets:
        wb.sheets(lang).delete()
    ws = wb.sheets.add(lang)
    ws.range("A1:Z1").rows.autofit()
    
    # DB 데이터 불러오기
    sql_col_set = db_tables(f"SELECT * FROM '중국어'")
    dataList = db_select(f"SELECT * FROM '{lang}'")
    
    # 컬럼 넓이 설정
    sql_col_list = db_columns(f"SELECT * FROM '{lang}'")
    testList = ["문자 넘침", "개행 오류", "다국어 기능과 의미 비매칭"]
    
    # 맨 첫줄 타이틀 쓰기
    for col_name in sql_col_list:
        ws[f'{chr(CHR_COL)}1'].value = col_name
        if col_name == "이미지":
            # ws.range(1, 1).column_width = 100
            ws.range(f'{chr(CHR_COL)}1').column_width = 100
        elif col_name in testList:
            ws.range(f'{chr(CHR_COL)}1').column_width = 50
        else:
            ws.range(f'{chr(CHR_COL)}1').column_width = 25
            
        CHR_COL += 1
    
    for i, data in enumerate(dataList):
        ws.range(f'A{i+2}').value=data
        ws.range(f'A{i+2}').row_height = 155
        img_path = data[0].replace("/", "\\")
        if not os.path.isfile(img_path):
            ws.pictures.add(img_path, 
                            left=ws.range(f"A{i+2}").left,
                            top=ws.range(f"A{i+2}").top,
                            width=400,
                            height=155)
        else:
            ws[f'A{i+2}'].value="파일 없음"
        
        

# wb.save(r"C:\Users\9350816\Desktop\다국어자동화(4).xlsx")
wb.save()
wb.close()





# https://blog.csdn.net/m0_64336020/article/details/121566766
# https://blog.csdn.net/m0_64336020/article/details/121587798