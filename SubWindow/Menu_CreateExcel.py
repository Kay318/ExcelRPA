"""
1. 큰 위젯기준으로 레이아웃2개 분리
1-2 지금까지 평가했던 평가항목들 스캔하기 그룹라디오
2. 분리된 레이아웃중 하나는 새로만드는것
3. 2번쟤꺼는 기존에 레이아웃에 추가하는것
"""

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from datetime import datetime
from Settings import Setup as sp
from functools import partial
import sys, os
from DataBase import DB as db
from progressBar import ProgressApp
from Log import LogManager
from Helper import *
import xlwings as xw

class UI_CreateExcel(QWidget):
    signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setupUI_CreateExcel()
        self.sp = sp.Settings()

    @AutomationFunctionDecorator
    def setupUI_CreateExcel(self):

        self.setFixedSize(430, 420)
        self.setWindowTitle("엑셀 생성")

        # 전체 화면 배치
        self.hBoxLayout = QHBoxLayout(self)
        self.vBoxLayout = QVBoxLayout(self)

        self.lang_vbox = QVBoxLayout()
        self.checkQ_label = QLabel()
        self.lang_widget = QWidget()
        self.lang_scroll =QScrollArea()
        self.lang_data_vbox = QVBoxLayout(self)
        self.checkQ_label.setText("언어 설정")
        self.lang_vbox.addWidget(self.checkQ_label)
        self.hBoxLayout.addLayout(self.lang_vbox)

        self.radioQGroupBox = QGroupBox("신규 및 기존 엑셀 생성")

        self.radio_vbox = QVBoxLayout()

        new_excel_label = QLabel()
        new_excel_label.setFixedHeight(80)
        font = QFont()
        font.setPixelSize(11)
        new_excel_label.setFont(font)
        new_excel_label.setAlignment(Qt.AlignTop)
        new_excel_label.setAlignment(Qt.AlignLeft)
        new_excel_label.setText(
        "\n\n저장한 데이터를 기반으로 엑셀을 새로 생성합니다.")
        set_excel_label = QLabel()
        set_excel_label.setFont(font)
        set_excel_label.setText(
        "\n\n기존 엑셀 기반으로 편집되어 작성됩니다.")
        set_excel_label.setFixedHeight(80)
        set_excel_label.setAlignment(Qt.AlignTop)
        set_excel_label.setAlignment(Qt.AlignLeft)

        self.radio_vbox.addWidget(new_excel_label)
        self.new_excel_groupBox = QGroupBox("신규 엑셀 생성")
        self.new_excel_groupBox.setCheckable(True)
        self.new_excel_groupBox.setFixedSize(275,50)
        self.new_excel_groupBox.clicked.connect(
            lambda : self.func_new_excel_groupBox())
        self.new_excel_func()
        self.radio_vbox.addWidget(self.new_excel_groupBox)

        self.radio_vbox.addWidget(set_excel_label)
        self.set_excel_groupBox = QGroupBox("기존 엑셀 편집")
        self.set_excel_groupBox.setFixedSize(275,50)
        self.set_excel_groupBox.setCheckable(True)
        self.set_excel_groupBox.clicked.connect(
            lambda : self.func_set_excel_groupBox())
        self.set_excel_func()
        self.radio_vbox.addWidget(self.set_excel_groupBox)

        self.new_excel_groupBox.setChecked(True)
        self.set_excel_groupBox.setChecked(False)

        self.radioQGroupBox.setLayout(self.radio_vbox)

        self.vBoxLayout.addWidget(self.radioQGroupBox)

        # 확인_취소
        self.create_excel = QWidget()
        self.create_excel.setFixedSize(292,80)
        check_cencel = QVBoxLayout()
        check = QPushButton("생성")
        check.clicked.connect(partial(self.func_check_run))
        cencel = QPushButton("취소")
        cencel.clicked.connect(partial(self.func_cancel))
        check_cencel.addWidget(check)
        check_cencel.addWidget(cencel)
        self.create_excel.setLayout(check_cencel)

        self.vBoxLayout.addWidget(self.create_excel)

        self.hBoxLayout.addLayout(self.vBoxLayout)
        self.langSetting()

    def langSetting(self):

        sql_tables = db.db_select("SELECT name FROM sqlite_master WHERE type='table';")
        dataList = [table[0] for table in sql_tables]

        self.lang_scroll.setFixedSize(110,380)
        self.lang_scroll.setWidgetResizable(True)
        self.lang_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.lang_data_vbox.setAlignment(Qt.AlignTop)
        the_check = QCheckBox("전체")

        if dataList != None and len(self.lang_data_vbox) != 0:

            item_list = list(range(self.lang_data_vbox.count()))
            item_list.reverse()#  Reverse delete , Avoid affecting the layout order 

            for i in item_list:

                item = self.lang_data_vbox.itemAt(i)
                self.lang_data_vbox.removeItem(item)
                if item.widget():
                    item.widget().deleteLater()

        self.lang_data_vbox.addWidget(the_check)

        for val in dataList:

            print(f"val {val}")
            globals()[f'checkBox_{val}'] = QCheckBox(val)
            globals()[f'checkBox_{val}'].clicked.connect(lambda : func_checkbox())
            self.lang_data_vbox.addWidget(globals()[f'checkBox_{val}'])

        the_check.clicked.connect(lambda : func_check())
        print(f"self.lang_data_vbox : {self.lang_data_vbox.count()}")
        self.lang_widget.setLayout(self.lang_data_vbox)
        self.lang_scroll.setWidget(self.lang_widget)
        self.lang_vbox.addWidget(self.lang_scroll)

        # 전체 체크박스 동작여부
        def func_check():

            if (the_check.isChecked() == True):
                for val in dataList:
                    globals()[f'checkBox_{val}'].setChecked(True)
            elif (the_check.isChecked() == False):
                for val in dataList:
                    globals()[f'checkBox_{val}'].setChecked(False)

        # 개별 체크박스 동작여부
        def func_checkbox():

            if (the_check.isChecked() == True):
                for val in dataList:

                    if (globals()[f'checkBox_{val}'].isChecked() == False):
                        the_check.setChecked(False)

            elif (the_check.isChecked() == False):

                result = []

                for val in dataList:

                    result.append(globals()[f'checkBox_{val}'].isChecked())

                if (result.count(False) == 0):
                    the_check.setChecked(True)
            
    @AutomationFunctionDecorator
    def new_excel_func(self):

        # 엑셀 경로 라벨
        path_hbox = QHBoxLayout()
        self.new_edit_path = QLineEdit()
        path_btn = QPushButton()
        path_btn.setMaximumWidth(30)
        path_btn.setText("...")

        path_btn.clicked.connect(partial(self.folder_toolButton_clicked, self.new_edit_path))
        path_hbox.addWidget(self.new_edit_path)
        path_hbox.addWidget(path_btn)

        self.new_excel_groupBox.setLayout(path_hbox)
        self.radio_vbox.addWidget(self.new_excel_groupBox)

    @AutomationFunctionDecorator
    def set_excel_func(self):

        # 엑셀 경로 라벨
        path_hbox = QHBoxLayout()
        self.set_edit_path = QLineEdit()
        path_btn = QPushButton()
        path_btn.setMaximumWidth(30)
        path_btn.setText("...")
        
        path_btn.clicked.connect(partial(self.langList_toolButton_clicked, self.set_edit_path))

        path_hbox.addWidget(self.set_edit_path)
        path_hbox.addWidget(path_btn)
        self.set_excel_groupBox.setLayout(path_hbox)
        self.radio_vbox.addWidget(self.set_excel_groupBox)

    @AutomationFunctionDecorator
    def langList_toolButton_clicked(self, edit, litter):
        """폴더 경로 불러오기

        Args:
            cnt: 변수명
        """
        folderPath = QFileDialog(self)
        folderPath.setFileMode(QFileDialog.FileMode.ExistingFile)
        folderPath.setNameFilter(self.tr("Data Files (*.csv *.xls *.xlsx);; All Files(*.*)"))
        folderPath.setViewMode(QFileDialog.ViewMode.Detail)

        if folderPath.exec_():
            fileNames = folderPath.selectedFiles()
            fileNames = str(fileNames)
            fileNames = fileNames[2 : len(fileNames) - 2]
            edit.setText(str(fileNames))
        else:
            edit.setText("")

    @AutomationFunctionDecorator
    def folder_toolButton_clicked(self, edit, litter):
        """폴더 경로 불러오기

        Args:
            cnt: 변수명
        """

        folderPath = QFileDialog.getExistingDirectory(self, 'Find Folder')

        if bool(folderPath):
            htime = datetime.now().strftime("%y%m%d_%H%M%S")
            fileName = f"{htime}_다국어자동화.xlsx"
            path = f'{folderPath}/{fileName}'
            edit.setText(path)

    @AutomationFunctionDecorator
    def func_new_excel_groupBox(self):

        if (self.new_excel_groupBox.isChecked() == True):
            self.new_excel_groupBox.setChecked(True)
            self.set_excel_groupBox.setChecked(False)
        else:
            self.new_excel_groupBox.setChecked(False)
            self.set_excel_groupBox.setChecked(True)

    @AutomationFunctionDecorator
    def func_set_excel_groupBox(self):

        if (self.set_excel_groupBox.isChecked() == True):
            self.set_excel_groupBox.setChecked(True)
            self.new_excel_groupBox.setChecked(False)
        else:
            self.set_excel_groupBox.setChecked(False)
            self.new_excel_groupBox.setChecked(True)
            
    def setting_Verification(self, langList):
        path = str(os.path.dirname(self.path))

        result = False
        
        if os.path.isdir(path):
            result = True
        else:
            btnReply = QMessageBox.warning(self, "주의", f"{path} 경로가 존재하지 않습니다.", QMessageBox.Ok, QMessageBox.Ok)
            LogManager.HLOG.info("엑셀 생성 팝업에서 존재하지 않는 경로 알림 표시")
            
            if btnReply == QMessageBox.Ok:
                result = False
                return result

        for lang in langList:
                
            path = os.path.dirname(self.imgCellCount(lang))
            if os.path.isdir(path):
                result = True
            else:
                btnReply = QMessageBox.warning(self, "주의", f"{self.imgCellCount(lang)[0]} 경로가 존재하지 않습니다.", QMessageBox.Ok, QMessageBox.Ok)
                LogManager.HLOG.info("언어 설정 팝업에서 존재하지 않는 경로 알림 표시")
                if btnReply == QMessageBox.Ok:
                    result = False
                    return result
    
        return result
    
    def imgCellCount(self, lang):
        """
        1. DB 받아오기
        2. 저장된 DB중 이미지 경로 있는것만 구별시키기
        3. 이미지 개수를 확인하고 : return 해당경로를 순서대로 List 반환
        """

        data = db.db_select_one(f"SELECT * FROM '{lang}'")

        return data[0]

    @AutomationFunctionDecorator
    def func_check_run(self, litter):
        self.testBool = False
        self.excel_groupBox_Bool = self.new_excel_groupBox.isChecked()
        for idx in range(len(self.lang_data_vbox)):
            
            if self.lang_data_vbox.itemAt(idx).widget().isChecked() == True:
                self.testBool = True
        
        if self.excel_groupBox_Bool:
            self.path = os.path.dirname(str(self.new_edit_path.text()))
        else:
            self.path = os.path.dirname(str(self.set_edit_path.text()))

        if not os.path.isdir(self.path):
            QMessageBox.warning(self, "주의", f"{self.path} 경로가 존재하지 않습니다.")
            LogManager.HLOG.info("엑셀 생성 팝업에서 존재하지 않는 경로 알림 표시")
            return
        
        self.lang_choice_list = []

        dataList, _ = self.sp.read_setup(table = "Language")
        
        for val in dataList:
            try:
                if (globals()[f'checkBox_{val}'].isChecked() == True):
                    self.lang_choice_list.append(val)
            except:
                pass

        for lang in self.lang_choice_list:
                
            path = os.path.dirname(self.imgCellCount(lang))
            if not os.path.isdir(path):
                QMessageBox.warning(self, "주의", f"{lang}에서 {self.imgCellCount(lang)} 경로가 존재하지 않습니다.", QMessageBox.Ok, QMessageBox.Ok)
                LogManager.HLOG.info("언어 설정 팝업에서 존재하지 않는 경로 알림 표시")
                return

        if self.testBool:
            if (self.new_excel_groupBox.isChecked() == True and self.new_edit_path.text() != ""):
                self.setEnabled(False)
                ProgressApp(self)
                LogManager.HLOG.info(f"엑셀 생성 누름 열림")
            elif (self.set_excel_groupBox.isChecked() == True and self.set_edit_path.text() != ""):
                try:
                    app = xw.App(visible=False, add_book=False)
                    app.display_alerts=False
                    wb = app.books.open(self.set_edit_path.text())
                    wb.sheets('SUMMARY')
                    wb.close()
                    app.quit()
                except:
                    QMessageBox.warning(self, '주의', '파일이 잘못되었습니다.\n재확인 해주세요.')
                    wb.close()
                    app.kill()
                    return
                    
                self.setEnabled(False)
                ProgressApp(self)
                LogManager.HLOG.info(f"엑셀 생성 누름 열림")
            else:
                QMessageBox.warning(self, '주의', '경로를 지정해주세요.')
        else:
            QMessageBox.warning(self, '주의', '언어를 선택해주세요.')

    @AutomationFunctionDecorator
    def func_cancel(self, litter):
        self.signal.emit()
        self.close()

    # 0726
    def closeEvent(self, litter) -> None:
        self.signal.emit()

    # 0725
    def keyPressEvent(self, a0: QKeyEvent) -> None:
        
        KEY_ENTER = 16777220
        KEY_CLOSE = 16777216

        print (f"a0.key() : {a0.key()}")
        if a0.key() == KEY_ENTER:
            self.func_check_run()
        elif a0.key() == KEY_CLOSE:
            self.close()

class ProgressBarThread(QThread):
    def __init__(self):
        super(ProgressBarThread, self).__init__()

    def run(self):
        self.progressbar = ProgressApp()

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    sys.exit(app.exec_())
