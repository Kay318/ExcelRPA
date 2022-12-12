import os
import glob
import sqlite3
from functools import wraps
# from multiprocessing import Process, Lock

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
        self.dbpath = "./DataBase/ExcelRPA.db"
        self.dbConn = sqlite3.connect(self.dbpath, isolation_level = None)
        self.c = self.dbConn.cursor()

    def close(self):
        self.dbConn.close()
        
def find_db():
    file = glob.glob("./DataBase/ExcelRPA.db")
    if file != []:
        return True
    else:
        return False

def remove_db():
    os.remove("./DataBase/ExcelRPA.db")

@DBDecorator
def db_edit(cmd: str):
    HDB.c.execute(cmd)

@DBDecorator
def db_insert(args1:str, args2:tuple):
    HDB.c.execute(args1, args2)

@DBDecorator
def db_select(cmd: str):
    HDB.c.execute(cmd)
    result = HDB.c.fetchall()
    return result

@DBDecorator
def db_select_one(cmd: str):
    HDB.c.execute(cmd)
    result = HDB.c.fetchone()
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
    
# def db_select(cmd: str):
#     HDB = DBManager()
#     try:
#         HDB.c.execute(cmd)
#     except:
#         msg = traceback.format_exc()
#         LogManager.HLOG.error(msg)
#         HDB.close()
#     HDB.close()

# lock = Lock()
# th_db = Process(target= db_control, args=(lock,))

if __name__ == "__main__":
    db_edit("DROP TABLE '중국어'")