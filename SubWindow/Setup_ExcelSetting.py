import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from pathlib import Path

sys.path.append(str(Path(__file__).parents[1]))
from Helper import *
from Log import LogManager
from Settings import Setup as sp
# from MainWindow import MainWindow as mainwindow

class Setup_ExcelSetting(QDialog):
    signal = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.sp = sp.Settings()
        self.setupUI_Excel_Setting()

    @AutomationFunctionDecorator
    def setupUI_Excel_Setting(self):
        self.setWindowTitle("엑셀 설정")

        self.setWindowFlags(
        Qt.Window |
        Qt.CustomizeWindowHint |
        Qt.WindowCloseButtonHint |
        Qt.WindowStaysOnTopHint
        )

        # 전체 화면 배치
        self.verticalLayout = QVBoxLayout(self)

        # 초기화 버튼
        self.reset_Button = QPushButton("초기화")
        self.verticalLayout.addWidget(self.reset_Button)

        self.img_cell_width_horizontal = QHBoxLayout()
        self.img_cell_width_label = QLabel("이미지 열 너비")
        self.img_cell_width_checkbox = QCheckBox("Auto")
        self.img_cell_width_lineEdit = QLineEdit()
        self.img_cell_width_horizontal.addWidget(self.img_cell_width_label)
        self.img_cell_width_horizontal.addWidget(self.img_cell_width_checkbox)
        self.img_cell_width_horizontal.addWidget(self.img_cell_width_lineEdit)
        self.verticalLayout.addLayout(self.img_cell_width_horizontal)

        self.start_settings = ["이미지 너비", "이미지 높이", "필드 너비", "평가 목록 너비"]
        self.start_settings_val = [15, 5, 40, 15]
        self.value_range = [("5~30"),("1~10"),("10~100"),("10~50")]
        dataList, _ = self.sp.read_setup(table = "Excel_Setting")

        auto_data = dataList.pop()

        self.img_cell_width_checkbox.setChecked(eval(auto_data))

        if eval(auto_data):
            self.img_cell_width_lineEdit.setEnabled(False)

        self.img_cell_width_lineEdit.setText(dataList.pop())

        for i in range(4):
            globals()[f'horizontalLayout{i}'] = QHBoxLayout()

            globals()[f'label{i}'] = QLabel()
            globals()[f'label{i}'].setText(f"{self.start_settings[i]}")
            globals()[f'horizontalLayout{i}'].addWidget(globals()[f'label{i}'])

            globals()[f'lineEdit{i}'] = QLineEdit()
            globals()[f'lineEdit{i}'].setText(dataList[i])
            globals()[f'lineEdit{i}'].setPlaceholderText(self.value_range[i])
            globals()[f'lineEdit{i}'].setFixedWidth(70)
            globals()[f'horizontalLayout{i}'].addWidget(globals()[f'lineEdit{i}'])
            self.verticalLayout.addLayout(globals()[f'horizontalLayout{i}'])

        # [확인], [취소] 버튼
        self.ok_horizontalLayout = QHBoxLayout()
        self.ok_horizontalLayout.setAlignment(Qt.AlignCenter)
        
        self.ok_Button = QPushButton("확인", self)
        self.ok_Button.setFocus()
        self.ok_horizontalLayout.addWidget(self.ok_Button)
        self.cancel_Button = QPushButton("취소", self)
        self.ok_horizontalLayout.addWidget(self.cancel_Button)
        self.verticalLayout.addLayout(self.ok_horizontalLayout)

        # 버튼 이벤트 함수
        self.tl_set_slot()

        self.tl_ini_set()

    def checkboxStateChanged(self):
        if self.img_cell_width_checkbox.isChecked() == True:
            self.img_cell_width_lineEdit.setEnabled(False)
        else:
            self.img_cell_width_lineEdit.setEnabled(True)

    @AutomationFunctionDecorator
    def tl_set_slot(self):
        self.reset_Button.clicked.connect(self.reset_Button_clicked)
        self.ok_Button.clicked.connect(self.ok_Button_clicked)
        self.cancel_Button.clicked.connect(self.close)
        self.img_cell_width_checkbox.stateChanged.connect(self.checkboxStateChanged)

    @AutomationFunctionDecorator
    def tl_ini_set(self):

        start_set = False
        for i in range(4):
            if globals()[f'lineEdit{i}'].text() == "":
                start_set = True
            else:
                start_set = False
                break
        
        if start_set:
            for i in range(4):
                globals()[f'lineEdit{i}'].setText(str(self.start_settings_val[i]))
            
    @AutomationFunctionDecorator
    def reset_Button_clicked(self, litter):
        LogManager.HLOG.info("엑셀 설정 팝업 확인 버튼 선택")
        reply = QMessageBox.question(self, '알림', '초기화 하시겠습니까?',
                                    QMessageBox.Ok | QMessageBox.No, QMessageBox.Ok)

        if reply == QMessageBox.Ok:
            LogManager.HLOG.info("엑셀 설정 초기화 선택")
            for i in range(4):
                globals()[f'lineEdit{i}'].setText(str(self.start_settings_val[i]))

        else:
            LogManager.HLOG.info("엑셀 설정 초기화 취소 선택")
            return

    @AutomationFunctionDecorator
    def ok_Button_clicked(self, litter):
        LogManager.HLOG.info("엑셀 설정 팝업 확인 버튼 선택")
        if self.img_cell_width_lineEdit.isEnabled() == True:
            if self.img_cell_width_lineEdit.text() != "":
                try:
                    int(self.img_cell_width_lineEdit.text())
                except:
                    QMessageBox.warning(self, '주의', f'{self.img_cell_width_label.text()} 수치를 정수형태로 지정해주세요.')
                    return

                if 26 > int(self.img_cell_width_lineEdit.text()) or int(self.img_cell_width_lineEdit.text()) > 80:
                    QMessageBox.warning(self, '주의', f'{self.img_cell_width_label.text()} 26에서 80 사이여야 합니다.')
                    return
            else:
                text1 = self.img_cell_width_lineEdit.text()
                QMessageBox.warning(self, '주의', f'{text1} 수치가 비어있습니다.')
                return

        # 중복 체크
        for i in range(4):
            if globals()[f'lineEdit{i}'].text() != "":

                text = globals()[f'label{i}'].text()
                try:
                    int(globals()[f'lineEdit{i}'].text())
                except:
                    QMessageBox.warning(self, '주의', f'{text} 수치를 정수형태로 지정해주세요.')
                    return

                if 5 > int(globals()[f'lineEdit{0}'].text()) or int(globals()[f'lineEdit{0}'].text()) > 30:
                    text = globals()[f'label{0}'].text()
                    QMessageBox.warning(self, '주의', f'{text} 5에서 30 사이여야 합니다.')
                    return
                elif 1 > int(globals()[f'lineEdit{1}'].text()) or int(globals()[f'lineEdit{1}'].text()) > 10:
                    text = globals()[f'label{1}'].text()
                    QMessageBox.warning(self, '주의', f'{text} 1에서 10 사이여야 합니다.')
                    return
                elif 10 > int(globals()[f'lineEdit{2}'].text()) or int(globals()[f'lineEdit{2}'].text()) > 100:
                    text = globals()[f'label{2}'].text()
                    QMessageBox.warning(self, '주의', f'{text} 10에서 100 사이여야 합니다.')
                    return
                elif 10 > int(globals()[f'lineEdit{3}'].text()) or int(globals()[f'lineEdit{3}'].text()) > 50:
                    text = globals()[f'label{3}'].text()
                    QMessageBox.warning(self, '주의', f'{text} 10에서 50 사이여야 합니다.')
                    return
            else:
                text = globals()[f'label{i}'].text()
                QMessageBox.warning(self, '주의', f'{text} 수치가 비어있습니다.')
                return

        self.sp.config["Excel_Setting"] = {}
        for i in range(4):
            self.sp.write_setup(table = "Excel_Setting", 
                                count=i, 
                                val=globals()[f'lineEdit{i}'].text(),
                                val2=None)
            LogManager.HLOG.info(f"{i+1}:엑셀 세팅 팝업에 {globals()[f'lineEdit{i}'].text()} 추가")

        self.sp.write_setup(table = "Excel_Setting", 
                            count=4, 
                            val=self.img_cell_width_lineEdit.text(),
                            val2=None)
        LogManager.HLOG.info(f"엑셀 세팅 팝업에 {self.img_cell_width_lineEdit.text()} 추가")

        self.sp.write_setup(table = "Excel_Setting", 
                            count=5, 
                            val=str(self.img_cell_width_checkbox.isChecked()),
                            val2=None)
        LogManager.HLOG.info(f"엑셀 세팅 팝업에 {str(self.img_cell_width_checkbox.isChecked())} 추가")

        self.signal.emit()
        self.destroy()

    def check_changedData(self):
        """변경사항이 있는지 체크하는 함수

        Returns:
            _type_: 변경사항이 있으면 True, 없으면 False
        """
        setupList, _ = self.sp.read_setup("Excel_Setting")
        lineList = [globals()[f'lineEdit{i}'].text() for i in range(4)]
        lineList.append(self.img_cell_width_lineEdit.text())
        lineList.append(str(self.img_cell_width_checkbox.isChecked()))
        
        if setupList != lineList:
            return True
        else:
            return False
        
    @AutomationFunctionDecorator
    def closeEvent(self, event) -> None:
        LogManager.HLOG.info("엑셀 설정 팝업 취소 버튼 선택")

        if self.check_changedData():
            reply = QMessageBox.question(self, '알림', '변경사항이 있습니다.\n취소하시겠습니까?',
                                    QMessageBox.Ok | QMessageBox.No, QMessageBox.No)

            if reply == QMessageBox.Ok:
                LogManager.HLOG.info("엑셀 설정 팝업 > 취소 > 변경사항 알림에서 예 선택")
                event.accept()
                self.signal.emit()
            else:
                LogManager.HLOG.info("엑셀 설정 팝업 > 취소 > 변경사항 알림에서 취소 선택")
                event.ignore()
        else:
            self.signal.emit()
                
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
    ui = Setup_ExcelSetting()
    ui.show()
    sys.exit(app.exec_())