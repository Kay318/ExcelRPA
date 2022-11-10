# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'setup_language.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

import os
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from functools import partial
from pathlib import Path

sys.path.append(str(Path(__file__).parents[1]))
from Helper import *
from Log import LogManager
from Settings import Setup as sp

class Setup_Language(QDialog):
    signal = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.sp = sp.Settings()
        self.setupUI_Language()

    @AutomationFunctionDecorator
    def setupUI_Language(self):
        # 해상도 받아옴
        screen = QDesktopWidget().screenGeometry()

        # 해상도에 따라 창 크기 설정
        main_width = screen.width() * 0.5                   # 메인창 넓이
        main_height = screen.height() * 0.5                 # 메인창 높이
        main_left = (screen.width() - main_width) / 2       # 메인창 x좌표
        main_top = (screen.height() - main_height) / 2      # 메인창 y좌표

        if main_width > 960:
            main_width = 960
        if main_height > 540:
            main_height = 540

        self.setMinimumSize(main_width, main_height)
        self.setGeometry(main_left, main_top, main_width, main_height)
        self.setWindowTitle("언어별 경로 설정")

        # 전체 화면 배치
        self.verticalLayout = QVBoxLayout(self)

        # [언어 추가], 언어 설정 리스트 영역 배치
        self.top_verticalLayout = QVBoxLayout()

        # [언어 추가] 버튼
        self.sl_editLang_horizontalLayout = QHBoxLayout()
        self.sl_editLang_horizontalLayout.setAlignment(Qt.AlignCenter)
        self.addLang_Button = QPushButton("언어 추가", self)
        self.addLang_Button.setMaximumWidth(130)
        self.sl_editLang_horizontalLayout.addWidget(self.addLang_Button)
        self.top_verticalLayout.addLayout(self.sl_editLang_horizontalLayout)

        # # 언어 설정 리스트 영역
        self.langList_scrollArea = QScrollArea(self)
        self.langList_scrollArea.setWidgetResizable(True)
        self.langList_scrollAreaWidgetContents = QWidget()
        self.langListScroll_verticalLayout = QVBoxLayout(self.langList_scrollAreaWidgetContents)
        self.langListScroll_verticalLayout.setAlignment(Qt.AlignTop)

        self.langList_scrollArea.setWidget(self.langList_scrollAreaWidgetContents)
        self.top_verticalLayout.addWidget(self.langList_scrollArea)
        self.verticalLayout.addLayout(self.top_verticalLayout)
        
        # [확인], [취소] 버튼
        self.sl_ok_horizontalLayout = QHBoxLayout()
        self.sl_ok_horizontalLayout.setAlignment(Qt.AlignRight)
        
        self.ok_Button = QPushButton("확인", self)
        self.sl_ok_horizontalLayout.addWidget(self.ok_Button)
        self.cancel_Button = QPushButton("취소", self)
        self.sl_ok_horizontalLayout.addWidget(self.cancel_Button)
        self.verticalLayout.addLayout(self.sl_ok_horizontalLayout)

        self.setLang_Button()
        self.sl_set_slot()

    @AutomationFunctionDecorator
    def sl_set_slot(self):
        self.addLang_Button.clicked.connect(partial(self.addLang_Button_clicked, lang="", path=""))
        self.ok_Button.clicked.connect(self.ok_Button_clicked)
        self.cancel_Button.clicked.connect(self.close)

    @AutomationFunctionDecorator
    def setLang_Button(self):
        """Setup.ini 파일에 데이터를 창에 표시
        """
        self.cnt = 0
        langList, langPath = self.sp.read_setup(table = "Language")
        LogManager.HLOG.info(f"{langList}, {langPath} 기준으로 창에 표시하기 시작")

        for lang, path in zip(langList, langPath):
            self.addLang_Button_clicked(lang, path)
        LogManager.HLOG.info(f"{langList}, {langPath} 기준으로 창에 표시 완료")

    def addLang_Button_clicked(self, lang="", path="", litter=None):
        """언어 추가 버튼 클릭

        Args:
            lang (str, optional): 언어명
            path (str, optional): 언어 경로
            litter (_type_, optional): _description_. Defaults to None.
        """
        lang_line_text = ""
        dir_line_text = ""
        
        if lang != "" :
            lang_line_text = lang
            dir_line_text = path

        globals()[f'langList_horizontalLayout{self.cnt}'] = QHBoxLayout()

        # 삭제 버튼
        globals()[f'del_langList_button{self.cnt}'] = QPushButton("-", self.langList_scrollAreaWidgetContents)
        globals()[f'del_langList_button{self.cnt}'].setMaximumWidth(30)
        globals()[f'del_langList_button{self.cnt}'].clicked.connect(partial(
            self.del_langList_button_clicked, layout = globals()[f'langList_horizontalLayout{self.cnt}']))
        globals()[f'langList_horizontalLayout{self.cnt}'].addWidget(globals()[f'del_langList_button{self.cnt}'])

        # 언어 입력
        globals()[f'lang_lineEdit{self.cnt}'] = QLineEdit(self.langList_scrollAreaWidgetContents)
        globals()[f'lang_lineEdit{self.cnt}'].setMaximumWidth(100)
        globals()[f'lang_lineEdit{self.cnt}'].setPlaceholderText('언어 입력')
        globals()[f'lang_lineEdit{self.cnt}'].setText(lang_line_text)
        globals()[f'lang_lineEdit{self.cnt}'].setFocus()
        globals()[f'langList_horizontalLayout{self.cnt}'].addWidget(globals()[f'lang_lineEdit{self.cnt}'])

        # 경로 입력
        globals()[f'dir_lineEdit{self.cnt}'] = QLineEdit(self.langList_scrollAreaWidgetContents)
        globals()[f'dir_lineEdit{self.cnt}'].setPlaceholderText('우측 버튼으로 폴더 경로 설정')
        globals()[f'dir_lineEdit{self.cnt}'].setText(dir_line_text)
        globals()[f'langList_horizontalLayout{self.cnt}'].addWidget(globals()[f'dir_lineEdit{self.cnt}'])

        # 경로 검색 버튼
        globals()[f'langList_toolButton{self.cnt}'] = QToolButton(self.langList_scrollAreaWidgetContents)
        globals()[f'langList_toolButton{self.cnt}'].setText("...")
        globals()[f'langList_horizontalLayout{self.cnt}'].addWidget(globals()[f'langList_toolButton{self.cnt}'])
        globals()[f'langList_toolButton{self.cnt}'].clicked.connect(partial(self.langList_toolButton_clicked, lineEdit = globals()[f'dir_lineEdit{self.cnt}']))
        self.langListScroll_verticalLayout.addLayout(globals()[f'langList_horizontalLayout{self.cnt}'])
        LogManager.HLOG.info(f"{globals()[f'lang_lineEdit{self.cnt}'].text()}, {globals()[f'dir_lineEdit{self.cnt}'].text()} 추가됨")
        self.cnt += 1

    def del_langList_button_clicked(self, layout):
        """라인 삭제 함수
        """

        cnt = 0
        for i in range(self.cnt):
            try:
                globals()[f'lang_lineEdit{i}'].text()
            except:
                cnt += 1

        if cnt + 1 == self.cnt:
            QMessageBox.warning(self, '주의', '삭제할 수 없습니다.\n최소 1가지 언어 설정은 하셔야 됩니다.')
            return

        for i in range(layout.count()):
            LogManager.HLOG.info(f"{layout.itemAt(i).widget().text()} 삭제 예정")
            layout.itemAt(i).widget().deleteLater()

    def langList_toolButton_clicked(self, lineEdit):
        """폴더 경로 불러오기

        Args:
            cnt: 변수명
        """
        folderPath = QFileDialog.getExistingDirectory(self, 'Select Folder')
        lineEdit.setText(folderPath)
        LogManager.HLOG.info(f"이미지 폴더:{lineEdit.text()} 설정")

    @AutomationFunctionDecorator
    def ok_Button_clicked(self, litter):
        """[확인] 버튼 클릭

        Args:
            litter (_type_): _description_
        """
        LogManager.HLOG.info("언어 설정 팝업 확인 버튼 선택")
        checkLang = []
        checkPath = []
        langPath = []

        # 빈칸 및 중복 언어 체크
        for i in range(self.cnt):
            try:
                if globals()[f'lang_lineEdit{i}'].text() in ["%", "'", "{", "}", ":", ";"]:
                    QMessageBox.warning(self, '주의', '["%", "\'", "\{", "\}", ":", ";"] 특수문자는 사용할 수 없습니다.')
                    LogManager.HLOG.info(f'언어 설정 팝업에서 특수문자 알림 표시')
                    return

                if len(globals()[f'lang_lineEdit{i}'].text()) > 10:
                    QMessageBox.warning(self, '주의', '언어명 최대 길이는 25자입니다.')
                    LogManager.HLOG.info(f'언어 설정 팝업에서 최대 길이 알림 표시')
                    return

                if globals()[f'lang_lineEdit{i}'].text() == "" or globals()[f'dir_lineEdit{i}'].text() == "":
                    QMessageBox.warning(self, '주의', '빈칸이 있습니다. \n 확인해 주세요.')
                    LogManager.HLOG.info("언어 설정 팝업에서 빈칸 알림 표시")
                    return
            except RuntimeError:
                continue

            if (globals()[f'lang_lineEdit{i}'].text() not in checkLang and
                globals()[f'dir_lineEdit{i}'].text() not in checkPath):
                checkLang.append(globals()[f'lang_lineEdit{i}'].text())
                checkPath.append(globals()[f'dir_lineEdit{i}'].text())
                langPath.append(((globals()[f'lang_lineEdit{i}'].text()), (globals()[f'dir_lineEdit{i}'].text())))
            else:
                QMessageBox.warning(self, '주의', '중복 라인이 있습니다.')
                LogManager.HLOG.info("언어 설정 팝업에서 중복 라인 알림 표시")
                return

            if os.path.isdir(globals()[f'dir_lineEdit{i}'].text()):
                pass
            else:
                QMessageBox.warning(self, "주의", f"{globals()[f'dir_lineEdit{i}'].text()} 존재하지 않습니다.")
                LogManager.HLOG.info("언어 설정 팝업에서 존재하지 않는 경로 알림 표시")
                return

        self.sp.config["Language"] = {}
        for i in range(self.cnt):
            try:
                self.sp.write_setup(table = "Language", 
                                    count=i, 
                                    val=globals()[f'lang_lineEdit{i}'].text(), 
                                    val2=globals()[f'dir_lineEdit{i}'].text())
                LogManager.HLOG.info(f"언어 설정 팝업에 {globals()[f'lang_lineEdit{i}'].text()}, {globals()[f'dir_lineEdit{i}'].text()} 추가")
            except RuntimeError:
                continue

        self.signal.emit(langPath)
        self.destroy()
        # QCoreApplication.instance().quit()

    @AutomationFunctionDecorator
    def closeEvent(self, event) -> None:
        LogManager.HLOG.info("언어 설정 팝업 취소 버튼 선택")
        setup_lang, setup_path  = self.sp.read_setup("Language")

        setupList = [data for data in zip(setup_lang, setup_path)]
        langList = []
        dirList = []
        
        for i in range(self.cnt):
            try:
                langList.append(globals()[f'lang_lineEdit{i}'].text())
                dirList.append(globals()[f'dir_lineEdit{i}'].text())
            except RuntimeError:
                continue

        lineList = [i for i in zip(langList, dirList)]

        if setupList != lineList:
            reply = QMessageBox.question(self, '알림', '변경사항이 있습니다.\n취소하시겠습니까?',
                                    QMessageBox.Ok | QMessageBox.No, QMessageBox.Ok)

            if reply == QMessageBox.Ok:
                LogManager.HLOG.info("언어 설정 팝업 > 취소 > 변경사항 알림에서 예 선택")
                event.accept()
                self.signal.emit([])
            else:
                LogManager.HLOG.info("언어 설정 팝업 > 취소 > 변경사항 알림에서 취소 선택")
                event.ignore()
        else:
            self.signal.emit([])

    @AutomationFunctionDecorator
    def keyPressEvent(self, a0: QKeyEvent) -> None:
        
        KEY_ENTER = 16777220
        KEY_SUB_ENTER = 16777221
        KEY_CLOSE = 16777216

        if a0.key() == KEY_ENTER or a0.key() == KEY_SUB_ENTER:
            self.ok_Button_clicked(None)
        elif a0.key() == KEY_CLOSE:
            self.close()

if __name__ == "__main__":
    LogManager.Init()
    app = QApplication(sys.argv)
    ui = Setup_Language()
    ui.show()
    sys.exit(app.exec_())
