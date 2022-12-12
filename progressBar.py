import sys
import os
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from SubWindow.excel_control import ExcelRun

class ProgressApp(QDialog):
    finished_excel_signal = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.is_done = 0
        self.loop = QEventLoop()
        QTimer.singleShot(99999999, self.loop.quit)
        self.path = parent.path
        self.selected_langList = parent.lang_choice_list
        self.testBool = parent.excel_groupBox_Bool
        if self.testBool:
            self.path_file = parent.new_edit_path.text()
        else:
            self.path_file = parent.set_edit_path.text()
        self.ce = parent
        self.initUI()
        self.run_excel_thread()
        self.show()
        self.loop.exec()

    def run_excel_thread(self):
        self.th_excel = ExcelRun(self)
        self.th_excel.progressBarValue.connect(self.callback)
        self.th_excel.signal_done.connect(self.callback_done)
        self.th_excel.start()

    def initUI(self):
        self.vbox = QVBoxLayout(self)
        self.hbox = QHBoxLayout(self)
        self.hbox.setAlignment(Qt.AlignCenter)
        self.pbar = QProgressBar(self)
        self.pbar.setValue(0)
        self.pbar.setRange(0,100)
        self.pbar.setGeometry(30, 40, 200, 25)
        self.pbar.setStyleSheet("QProgressBar { border: 2px solid grey; border-radius: 5px; color: rgb(20,20,20);  background-color: #FFFFFF; text-align: center;}QProgressBar::chunk {background-color: rgb(100,200,200); border-radius: 10px; margin: 0.1px;  width: 1px;}")
        self.vbox.addWidget(self.pbar)
        self.pbar.is_done = 0

        self.btn = QPushButton('취소', self)
        self.btn.setFixedWidth(100)
        self.btn.clicked.connect(self.close)
        self.hbox.addWidget(self.btn)
        self.vbox.addLayout(self.hbox)

    def callback(self, i):
        self.pbar.setValue(i)
        
    def callback_done(self, i):
        self.is_done = i
        if self.is_done == 1:
            self.close()

    def closeEvent(self, event) -> None:
        if self.is_done == 1:
            path = str(os.path.dirname(self.path_file))
            os.startfile(path)
            self.ce.close()
            self.close()
        else:
            quit_msg = "엑셀 생성을 종료하시겠습니까?"
            reply = QMessageBox.question(self, 'Message', quit_msg, QMessageBox.Yes, QMessageBox.No)
            if reply == QMessageBox.Yes:
            # 멀티쓰레드를 종료하는 stop 메소드를 실행함
                self.th_excel.stop()
                event.accept()
                self.ce.setEnabled(True)
                self.close()
            else:
                event.ignore()

    
        # self.th_excel = ExcelRun(self)
        # self.th_excel.progressBarValue.connect(self.callback)
        # self.th_excel.signal_done.connect(self.callback_done)
        # self.th_excel.start()
        
        # QApplication.processEvents()
        
        # self.btn = QPushButton('Cancel', self)
        # self.btn.move(40, 80)
        # self.btn.clicked.connect(self.doAction)

        # self.timer = QBasicTimer()
        # self.step = 0

        # self.setWindowTitle('진행도')
        # self.setGeometry(300, 300, 300, 200)
        
    
# class ProgressApp(QDialog):

#     def __init__(self, time, new_set_difference, save_path, wb:object):
#         super().__init__()
#         self.difference = new_set_difference
#         self.path = save_path
#         self.wb = wb
#         self.save_Bool = False

#         loop = QEventLoop()
#         QTimer.singleShot(1000, loop.quit)

#         self.numberVar = time

#         self.setWindowTitle('진행도')
#         self.setFixedSize(230, 150)
#         self.pbar = QProgressBar(self)
#         self.pbar.setGeometry(30, 40, 200, 25)
    
#         self.btn = QPushButton('Stop', self)
#         self.btn.move(40, 80)
#         self.btn.clicked.connect(self.doAction)

#         self.timer = QBasicTimer()
#         self.step = 0
#         self.show()
#         self.timer.start(self.numberVar, self)
#         loop.exec_()

#     def timerEvent(self, e):
#         if self.step == 99 and self.difference:
#             self.timer.stop()

#             idx = 1
#             while(os.path.isfile(self.path)):
#                 self.path = str(os.path.dirname(self.path))
#                 self.path = f"{self.path}\\다국어평가결과({idx}).xlsx"
#                 idx = idx + 1
#             else:
#                 self.wb.save(self.path)
#                 QApplication.processEvents()
#                 self.timer.start(self.numberVar, self)
            
#         elif self.step >= 100:
#             self.timer.stop()
#             self.btn.setText('Finished')
#             return 

#         self.step = self.step + 1
#         self.pbar.setValue(self.step)

#     def doAction(self):
#         if self.timer.isActive():
#             self.timer.stop()
#             self.btn.setText('Start')
#         elif self.btn.text() == "Finished":
#             self.save_Bool = True
#             self.close()
#         else:
#             reply = QMessageBox.question(self, '알림', '현재 엑셀 작업을 취소하시겠습니까?',
#                                         QMessageBox.Ok | QMessageBox.No | QMessageBox.Cancel, QMessageBox.Ok)
#             if reply == QMessageBox.Ok:
#                 self.close()
#             else:
#                 self.timer.start(self.numberVar, self)
#                 self.btn.setText('Stop')
#         QApplication.processEvents()

#     def closeEvent(self, a0) -> None:
        
#         path = str(os.path.dirname(self.path))
#         os.startfile(path)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ProgressApp()
    ex.show()
    sys.exit(app.exec_())