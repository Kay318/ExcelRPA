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

class Setup_TestList(QDialog):
    signal = pyqtSignal(list)
    def __init__(self, parent=None):
        super().__init__(parent)
        self.fieldList = parent.fieldList
        self.sp = sp.Settings()
        self.setupUI_TestList()

    @AutomationFunctionDecorator
    def setupUI_TestList(self):
        self.setWindowTitle("평가목록 설정")

        # 전체 화면 배치
        self.verticalLayout = QVBoxLayout(self)

        # Setup.ini 파일에 데이터를 창에 표시
        testList, _ = self.sp.read_setup(table = "Test_List")
        check_first = True

        for i in range(8):
            globals()[f'horizontalLayout{i}'] = QHBoxLayout()

            globals()[f'label{i}'] = QLabel()
            globals()[f'label{i}'].setText(f"{i+1}")
            globals()[f'horizontalLayout{i}'].addWidget(globals()[f'label{i}'])

            globals()[f'lineEdit{i}'] = QLineEdit()
            globals()[f'horizontalLayout{i}'].addWidget(globals()[f'lineEdit{i}'])
            self.verticalLayout.addLayout(globals()[f'horizontalLayout{i}'])
            try:
                globals()[f'lineEdit{i}'].setText(testList[i])
            except:
                pass

            # 포커스 설정: 빈칸 혹은 마지막칸
            if (globals()[f'lineEdit{i}'].text() == "" or i==7) and check_first:
                globals()[f'lineEdit{i}'].setFocus()
                check_first = False

        # [확인], [취소] 버튼
        self.ok_horizontalLayout = QHBoxLayout()
        self.ok_horizontalLayout.setAlignment(Qt.AlignRight)
        
        self.ok_Button = QPushButton("확인", self)
        self.ok_horizontalLayout.addWidget(self.ok_Button)
        self.cancel_Button = QPushButton("취소", self)
        self.ok_horizontalLayout.addWidget(self.cancel_Button)
        self.verticalLayout.addLayout(self.ok_horizontalLayout)

        # 버튼 이벤트 함수
        self.tl_set_slot()

    @AutomationFunctionDecorator
    def tl_set_slot(self):
        self.ok_Button.clicked.connect(self.ok_Button_clicked)
        self.cancel_Button.clicked.connect(self.close)

    @AutomationFunctionDecorator
    def ok_Button_clicked(self, litter):
        LogManager.HLOG.info("평가 목록 설정 팝업 확인 버튼 선택")
        testList = []

        # 중복 체크
        for i in range(8):
            if globals()[f'lineEdit{i}'].text() != "":
                if globals()[f'lineEdit{i}'].text() not in testList:
                    testList.append(globals()[f'lineEdit{i}'].text())
                else:
                    QMessageBox.warning(self, '주의', '중복 라인이 있습니다.')
                    LogManager.HLOG.info("평가 목록 팝업에서 중복 라인 알림 표시")
                    return

                if globals()[f'lineEdit{i}'].text() in self.fieldList:
                    x = globals()[f'lineEdit{i}'].text()
                    QMessageBox.warning(self, '주의', f'"{x}"는 필드에도 있습니다.')
                    LogManager.HLOG.info(f'평가 목록 팝업과 필드 설정 팝업에서 "{x}" 겹침 알림 표시')
                    return

        self.sp.config["Test_List"] = {}
        for i in range(8):
            if globals()[f'lineEdit{i}'].text() != "":
                self.sp.write_setup(table = "Test_List", 
                                    count=i, 
                                    val=globals()[f'lineEdit{i}'].text(),
                                    val2=None)
                LogManager.HLOG.info(f"{i+1}:평가 목록 팝업에 {globals()[f'lineEdit{i}'].text()} 추가")
        if testList == []:
            testList = ["OK"]
            self.sp.clear_table("Test_List")
        self.signal.emit(testList)
        self.destroy()
        # QCoreApplication.instance().quit()
        
    @AutomationFunctionDecorator
    def closeEvent(self, event) -> None:
        LogManager.HLOG.info("필드 설정 팝업 취소 버튼 선택")
        setupList, _ = self.sp.read_setup("Test_List")
        lineList = [globals()[f'lineEdit{i}'].text() for i in range(8) if globals()[f'lineEdit{i}'].text() != ""]

        if setupList != lineList:
            reply = QMessageBox.question(self, '알림', '변경사항이 있습니다.\n취소하시겠습니까?',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

            if reply == QMessageBox.Yes:
                LogManager.HLOG.info("필드 설정 팝업 > 취소 > 변경사항 알림에서 예 선택")
                event.accept()
                self.signal.emit([])
            else:
                LogManager.HLOG.info("필드 설정 팝업 > 취소 > 변경사항 알림에서 취소 선택")
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
    ui = Setup_TestList()
    ui.show()
    sys.exit(app.exec_())