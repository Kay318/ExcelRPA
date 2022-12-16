from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

class LoadingScreen(QWidget):
    
    def __init__(self, parent):
        super().__init__(parent)    
        ph = self.parent().geometry().height()
        pw = self.parent().geometry().width()
        self.setFixedSize(pw, ph) 
        size = self.size()
        self.move(int(pw/2 - size.width()/2), int(ph/2 - size.height()/2))
        # self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.CustomizeWindowHint)
        self.setWindowFlags(Qt.FramelessWindowHint)

        self.label_animation = QLabel(self)
        self.label_animation.resize(pw, ph)
        self.movie = QMovie('./IMG_Source/loading1.gif')
        self.label_animation.setMovie(self.movie)
        self.label_animation.setAlignment(Qt.AlignCenter)

    def startAnimation(self):
        opacity_effect = QGraphicsOpacityEffect(self)
        opacity_effect.setOpacity(0.1)
        self.movie.start()
        self.show()

    def stopAnimation(self):
        self.movie.stop()
        self.close()