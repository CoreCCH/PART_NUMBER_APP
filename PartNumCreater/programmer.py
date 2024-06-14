#GUI imports
import sys
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton, QVBoxLayout, QWidget, QFileDialog, QGridLayout
from PyQt5.QtGui import QPixmap
from PyQt5 import QtGui, QtCore
from PyQt5.QtGui import QCursor
#function imports
from functions import frame1, frame2, grid

#initiallize GUI application
app = QApplication(sys.argv)

for i in range(10):
    for j in range(8):
        label = QLabel()
        if(j == 5 or j == 0 or j == 7):
            label.setFixedWidth(60)  # 设置每个标签的固定大小为 100x100 像素
        elif(i == 0 or i == 9):
            label.setFixedHeight(30)
        elif(j == 1 or j == 3):
            label.setFixedWidth(100)
        else:
            label.setFixedSize(210, 80)  # 设置每个标签的固定大小为 100x100 像素
        grid.addWidget(label, i, j, 1, 1)

#window and settings
window = QWidget()
window.setWindowTitle("料號產生器")
# window.setFixedWidth(1000)
#place window in (x,y) coordinates
# window.move(2700, 200)
window.setStyleSheet("background: #F5F5F5;")
window.resize(1350, 800)

#display frame 1
frame1()

window.setLayout(grid)
grid.setSpacing(0)

window.show()
sys.exit(app.exec()) #terminate the app
