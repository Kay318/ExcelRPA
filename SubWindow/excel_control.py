import xlwings as xw
from datetime import datetime
from Settings import Setup as sp
import os
import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from pathlib import Path
from Log import LogManager

sys.path.append(str(Path(__file__).parents[1]))
from DataBase import DB as db

class ExcelRun(QThread):
    progressBarValue = pyqtSignal(int)
    signal_done = pyqtSignal(int)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.path = parent.path
        self.lang_List = parent.selected_langList
        self.testBool = parent.testBool
        self.path_file = parent.path_file
        self.pre_ws = None
        
        self.set_row = 1
        self.start_column = 1
        self.lang_cnt = 1
        self.sp = sp.Settings()
        self.testList, _ = self.sp.read_setup(table = "Test_List")
        self.excel_setup()
        
    def run(self):
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts=False

        if self.testBool:
            try:
                wb = self.app.books.add()
                wb.sheets[0].name = "SUMMARY"
                self.ws_summary = wb.sheets('SUMMARY')
                
                for idx, lang in enumerate(self.lang_List):
                    self.set_percent(idx)
                    
                    # 새로 생성하는 시트가 맨뒤에서 생성됨
                    if bool(self.pre_ws):
                        self.ws = wb.sheets.add(lang, after=self.pre_ws)
                    else:
                        self.ws = wb.sheets.add(lang)

                    self.set_langSheet(lang)

                    self.ROW_NUM = 1
                    
                    # SUMMARY시트 입력
                    while True:
                        summary_languageList = self.ws_summary.range(f"A{self.ROW_NUM}").expand("down").value
                        
                        if not bool(summary_languageList):
                            self.insert_summary(lang)
                            break
                                
                        self.ROW_NUM = self.ROW_NUM + len(summary_languageList) + 1
                        
                    self.pre_ws = self.ws
                        
                # SUMMARY 시트를 맨앞으로 당겨감
                sheets = wb.sheets
                self.ws_summary.api.Move(Before=sheets[0].api, After=None)
                
                wb.save(self.path_file)
                self.progressBarValue.emit(100)

            except Exception as e:
                LogManager.HLOG.info(f"엑셀 생성 중 오류", e)
                QMessageBox.warning(self.parent, '주의', '엑셀 생성이 실패되었습니다.\n파일 끄고 다시 해주세요.')
            finally:
                wb.close()
                self.app.quit()
                self.signal_done.emit(1)
                
        else:
            try:
                wb = self.app.books.open(self.path_file)
                self.ws_summary = wb.sheets('SUMMARY')
                sheets = wb.sheets
                sheets_li = [s.name for s in sheets]
                
                for idx, lang in enumerate(self.lang_List):
                    self.set_percent(idx)

                    # 시트 있으면 삭제 후 생성
                    if lang in sheets_li:
                        wb.sheets(lang).delete()
                    self.ws = wb.sheets.add(lang)
                    
                    self.set_langSheet(lang)
                    
                    self.ROW_NUM = 1
                    
                    # SUMMARY시트 입력
                    while True:
                        summary_languageList = self.ws_summary.range(f"A{self.ROW_NUM}").expand("down").value
                        if not bool(summary_languageList):
                            self.insert_summary(lang)
                            break
                        
                        if lang in summary_languageList:
                            self.ws_summary.range(f'{self.ROW_NUM}:{self.ROW_NUM+len(summary_languageList)}').delete()
                            self.insert_summary(lang)
                            break
                                
                        self.ROW_NUM = self.ROW_NUM + len(summary_languageList) + 1

                    # 기존 엑셀의 시트 순서대로 배치, 없는 시트는 맨마지막에 배치
                    if lang in sheets_li:
                        self.ws.api.Move(Before=None, After=sheets[sheets_li.index(lang)].api)
                    else:
                        self.ws.api.Move(Before=None, After=sheets[len(sheets_li)].api)
                    
                # SUMMARY 시트를 맨앞으로 당겨감
                self.ws_summary.api.Move(Before=sheets[0].api, After=None)
                
                htime = datetime.now().strftime("%y%m%d_%H%M%S")
                fileName = f"{htime}_다국어자동화.xlsx"
                savePath = os.path.join(self.path, fileName)
                wb.save(savePath)
                self.progressBarValue.emit(100)
                
            except Exception as e:
                LogManager.HLOG.info(f"기존 엑셀 편집 중 오류", e)
                QMessageBox.warning(self.parent, '주의', '엑셀 편집이 실패되었습니다.\n파일 끄고 다시 해주세요.')
            finally:
                wb.close()
                self.app.quit()
                self.signal_done.emit(1)
            
    def set_langSheet(self, lang):
        """언어별 시트의 모든 내용

        Args:
            lang (_type_): 현재 진행중인 언어
        """
        # DB에서 필요한 데이터 불러오기
        self.select_DB(lang)

        self.insert_FirstLine()
        self.insert_langSheetData()
        self.set_langSheetStyle()
            
    def set_percent(self, idx):
        """ProgressBar 계산을 위해 필요한 수치 정의
           self.start_percent: 세부적으로 시작하는 퍼센트
           self.split_percent: 언어의 갯수에 따라 N등분 한 퍼센트

        Args:
            idx (_type_): 현재 진행중인 언어 index
        """
        self.start_percent = idx/len(self.lang_List)
        self.split_percent = 1/len(self.lang_List)

    def insert_langSheetData(self):
        """언어별 시트에 데이터 입력하는 함수
        """

        self.summaryData = []

        LogManager.HLOG.info(f"이미지 너비:",  self.IMG_WIDTHSIZE)
        LogManager.HLOG.info(f"이미지 높이:",  self.IMG_HEIGHTSIZE)
        LogManager.HLOG.info(f"필드 너비:",  self.SHEET_WIDTHSIZE)
        LogManager.HLOG.info(f"평가 목록 너비:",  self.SHEET_EvaluationListSIZE)
        LogManager.HLOG.info(f"이미지셀 너비:",  self.IMG_FAINAL_WIDTH)

        for i, data in enumerate(self.dataList):
            # 데이터 입력
            self.ws.range(f'A{i+2}').value=data

            # 이미지 삽입
            img_path = data[0].replace("/", "\\")
            if os.path.isfile(img_path):
                self.ws.pictures.add(img_path, 
                                left=self.ws.range(f"A{i+2}").left,
                                top=self.ws.range(f"A{i+2}").top,
                                width=self.IMG_WIDTHSIZE,
                                height=self.IMG_HEIGHTSIZE)
            else:
                self.ws[f'A{i+2}'].value=f"파일 없음\n{img_path}"

            # 이미지 셀 너비, 높이 설정
            self.ws.range(f'A{i+2}').row_height = self.IMG_HEIGHTSIZE
            self.ws.range(f'A{i+2}').column_width = self.IMG_FAINAL_WIDTH
                       
            # SUMMARY시트에 삽입할 데이터 저장
            if data[1:len(self.testList)+1].count('PASS') != len(self.testList):
                self.summaryData.append(data)

            percent_val = round((self.start_percent + ((i+1)/len(self.dataList))*self.split_percent)*100)
            if percent_val > 99:
                percent_val = 99
            self.progressBarValue.emit(percent_val)

    def select_DB(self, lang):
        """DB 조회하면서 컬럼명, 내용을 세팅하는 함수

        Args:
            lang (_type_): 현재 편집중인 언어
        """
        self.dataList = db.db_select(f"SELECT * FROM '{lang}'")
        self.sql_col_list = db.db_columns(f"SELECT * FROM '{lang}'")
        self.summary_col_list = self.sql_col_list[:]
        self.summary_col_list.insert(0, '언어')

    def insert_FirstLine(self):
        """언어별 타이틀 쓰는 함수
        """

        for i, col_name in enumerate(self.sql_col_list):
            self.ws.cells(1, i+1).value = col_name
            if col_name in self.testList:
                self.ws.cells(1, i+1).column_width = self.SHEET_EvaluationListSIZE
            else:
                self.ws.cells(1, i+1).column_width = self.SHEET_WIDTHSIZE
                
        firstRange = self.ws.range("A1").expand('right')
        firstRange.rows.autofit()
        firstRange.api.WrapText = True
        firstRange.color = xw.utils.rgb_to_int((153,204,000))       # 배경색: 초록색

    def setBorder(self, range):
        """전체 테두리 추가

        Args:
            range (_type_): 적용할 range
        """
        range.api.Borders.Weight = 2
        
    def alignCenter(self, range):
        """가운데 맞춤

        Args:
            range (_type_): 적용할 range
        """
        range.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        range.api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter

    def stop(self):
        self.app.kill()
        self.terminate()

    def insert_summary(self, lang):
        """SUMMARY 입력하는 함수

        Args:
            lang (_type_): 편집 중인 언어
        """
        if bool(self.summaryData):
            START_ROW = self.ROW_NUM
            self.ws_summary.range(f'{self.ROW_NUM}:{self.ROW_NUM+len(self.summaryData)+1}').insert()
            
            self.insert_summaryTitle()
            self.ROW_NUM += 1
            self.insert_summaryData(lang)
            self.set_summaryStyle(START_ROW)

    def insert_summaryTitle(self):
        """SUMMARY 타이틀 입력하는 함수
           : 가운데 맞춤, 배경색, 자동 줄바꿈, 행높이 자동 설정 포함
        """
        for i, title in enumerate(self.summary_col_list):
            self.ws_summary.cells(self.ROW_NUM, i+1).value = title
            if title in self.testList:
                self.ws_summary.cells(self.ROW_NUM, i+1).column_width = self.SHEET_EvaluationListSIZE
            elif title == "언어":
                self.ws_summary.cells(self.ROW_NUM, i+1).column_width = 14
            else:
                self.ws_summary.cells(self.ROW_NUM, i+1).column_width = self.SHEET_WIDTHSIZE
                
        titleRange = self.ws_summary.range(f"A{self.ROW_NUM}").expand('right')
        titleRange.rows.autofit()
        titleRange.api.WrapText = True
        titleRange.color = xw.utils.rgb_to_int((153,204,000))       # 배경색: 초록색

    def insert_summaryData(self, lang):
        """SUMMARY 시트에서 선택한 언어에 대한 데이터 입력하는 함수

        Args:
            lang (_type_): 선택한 언어
        """
        for data in self.summaryData:
            self.ws_summary.cells(self.ROW_NUM, 1).value = lang
            self.ws_summary.range(f"B{self.ROW_NUM}").value = data            
            self.ROW_NUM += 1

    def set_summaryStyle(self, START_ROW):
        """SUMMARY시트에서 해당 언어 영역 스타일(테두리, 자동줄바꿈, 가운데 맞춤) 설정

        Args:
            START_ROW (_type_): 해당 언어가 시작되는 행
        """
        summaryTableRange = self.ws_summary.range(f"A{START_ROW}").expand('table')
        self.set_style(summaryTableRange)

    def set_langSheetStyle(self):
        """언어별 시트에서 스타일(테두리, 자동줄바꿈, 가운데 맞춤) 설정
        """
        tableRange = self.ws.range("A1").expand('table')
        self.set_style(tableRange)

    def set_style(self, range):
        """테두리, 자동줄바꿈, 가운데 맞춤
        """
        self.alignCenter(range)
        self.setBorder(range)
        range.api.WrapText = True

    def excel_setup(self):
        """이미지 사이즈, 열너비를 정의하는 함수
        """
        excel_setList, _ = self.sp.read_setup(table = "Excel_Setting")

        self.IMG_WIDTHSIZE = int(excel_setList[0]) * 15 / 0.53      # 이미지 너비
        self.IMG_HEIGHTSIZE = int(excel_setList[1]) * 15 / 0.53     # 이미지 높이
        self.SHEET_WIDTHSIZE = int(excel_setList[2])                # 필드 너비
        self.SHEET_EvaluationListSIZE = int(excel_setList[3])       # 평가 목록 너비
        
        if excel_setList[5] == 'False':
            self.IMG_FAINAL_WIDTH = int(excel_setList[4])
        else:
            self.IMG_FAINAL_WIDTH = self.IMG_WIDTHSIZE * 70.25 / 425    # 이미지 시트 너비