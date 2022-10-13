import os
import sqlite3

class DBManager:
    def __init__(self):
        self.dbpath = f"{os.getcwd()}\\Database\\ExcelRPA.db"
        self.dbConn = sqlite3.connect(self.dbpath)
        self.c = self.dbConn.cursor()

    def close(self):
        self.dbConn.close()

    def create_target(self, TEXT):
        self.dbConn.executescript(
            TEXT
        )

if __name__ == "__main__":
    db = DBManager()
    db.close()