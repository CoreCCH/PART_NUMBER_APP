#GUI imports
import sys
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton, QVBoxLayout, QWidget, QFileDialog, QGridLayout
from PyQt5.QtGui import QPixmap
from PyQt5 import QtGui, QtCore
from PyQt5.QtGui import QCursor
#function imports
from part_code_generator import frame1
from part_code_searcher import frame_search_page
from barcode_generator import barcode_select_page
from main_page import main_page
from component import grid

#initiallize GUI application
app = QApplication(sys.argv)

for i in range(10):
    for j in range(8):
        label = QLabel()
        # label.setStyleSheet("background: white;")
        if(j == 5 or j == 0 or j == 7):
            label.setFixedWidth(60)  
        elif(i == 0 or i == 9):
            label.setFixedHeight(30)
        elif(j == 1 or j == 3):
            label.setFixedWidth(150)
        else:
            label.setFixedSize(150, 80)  
        grid.addWidget(label, i, j, 1, 1)

#window and settings
window = QWidget()
window.setWindowTitle("EPR輔助軟體")
# window.setFixedWidth(1000)
#place window in (x,y) coordinates
# window.move(2700, 200)
window.setStyleSheet("background: #F5F5F5;")
window.resize(1350, 800)

#display frame 1
barcode_select_page()
# main_page()

window.setLayout(grid)
grid.setSpacing(0)

window.show()
sys.exit(app.exec()) #terminate the app
