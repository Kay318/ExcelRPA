
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
        self.dbpath = "D:/Skillup/new/DataBase/ExcelRPA.db"
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

def insertFirstLine(ws2, sql_col_list):
    """첫줄 타이틀 쓰는 함수
    """

    for i, col_name in enumerate(sql_col_list):
        ws2.cells(1, i+1).value = col_name
        if col_name in testList:
            ws2.cells(1, i+1).column_width = 15
        else:
            ws2.cells(1, i+1).column_width = 40
            
    firstRange = ws2.range("A1").expand('right')
    firstRange.rows.autofit()
    firstRange.api.WrapText = True
    firstRange.color = xw.utils.rgb_to_int((153,204,000))       # 배경색: 초록색
    
def setBorder(range):
    """전체 테두리 추가

    Args:
        range (_type_): 적용할 range
    """
    range.api.Borders.Weight = 2
    
def alignCenter(range):
    """가운데 맞춤

    Args:
        range (_type_): 적용할 range
    """
    range.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    range.api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
    

path = r"C:\Users\9350816\Desktop\다국어자동화(7).xlsx"
 
app = xw.App()
wb = app.books.open(path)
print(wb.name)

# 모든 시트명 표시
sheets = wb.sheets
print(sheets)
li_sheets = [s.name for s in sheets]
print(li_sheets)

langList = ["중국어"]


for idx, lang in enumerate(langList):
    if lang in li_sheets:
        wb.sheets(lang).delete()
    ws2 = wb.sheets.add(lang)
    print(ws2)
    print(sheets[1])
    ws2.api.Move(Before=sheets[1].api, After=None)

    print(ws2.index)
    
    # DB 데이터 불러오기
    dataList = db_select(f"SELECT * FROM '{lang}'")
    sql_col_list = db_columns(f"SELECT * FROM '{lang}'")
    testList = ['문자 넘침', '개행 오류', '다국어 기능과 의미 비매칭', '축약어']

    # print(f"IMG_WIDTHSIZE: {self.IMG_WIDTHSIZE}")
    # print(f"IMG_HEIGHTSIZE: {self.IMG_HEIGHTSIZE}")
    # print(f"IMG_FAINAL_WIDTH: {self.IMG_FAINAL_WIDTH}")
    # print(f"IMG_SHEET_HEIGHTSIZE: {self.IMG_SHEET_HEIGHTSIZE}")
    
    # 맨 첫줄 타이틀 쓰기
    insertFirstLine(ws2, sql_col_list)
    
    summaryData = []
    for i, data in enumerate(dataList):
        # 데이터 입력
        ws2.range(f'A{i+2}').value=data

        # 이미지 삽입
        img_path = data[0].replace("/", "\\")
        if os.path.isfile(img_path):
            ws2.pictures.add(img_path, 
                            left=ws2.range(f"A{i+2}").left,
                            top=ws2.range(f"A{i+2}").top,
                            width=424,
                            height=141)
        else:
            ws2[f'A{i+2}'].value=f"파일 없음\n{img_path}"
        
        # 이미지 셀 너비, 높이 설정
        ws2.range(f'A{i+2}').row_height = 141
        ws2.range(f'A{i+2}').column_width = 70

        # SUMMARY시트에 삽입할 데이터 저장
        if data[1:len(testList)+1].count('PASS') != len(testList):
            summaryData.append(data)

    tableRange = ws2.range("A1").expand('table')
    alignCenter(tableRange)
    setBorder(tableRange)

    ws_summary = wb.sheets('SUMMARY')
    summary_languageRange = ws_summary.range("A21").expand("down")
    # summary_firstLineRange = ws_summary.range("A1").expand('right')
    # print(summary_languageRange)
    print(summary_languageRange.value)
    print(ws_summary['A1'].end('down').row)
    
    ROW_NUM = 1
    while True:
        summary_languageList = ws_summary.range(f"A{ROW_NUM}").expand("down").value
        if not bool(summary_languageList):
            break
        
        if '독일어' in summary_languageList:
            ws_summary.range(f'{ROW_NUM}:{ROW_NUM+len(summary_languageList)}').delete()
            # ws_summary.range('6:7').delete()
        
        ROW_NUM = ROW_NUM + len(summary_languageList) + 1
            
        
        
    
    
        
        

# wb.save(r"C:\Users\9350816\Desktop\다국어자동화(4).xlsx")
# wb.save()
# wb.close() 







# https://blog.csdn.net/m0_64336020/article/details/121566766
# https://blog.csdn.net/m0_64336020/article/details/121587798