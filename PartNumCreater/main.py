#GUI imports pyinstaller --add-binary "libusb-1.0.dll:." -F PartNumCreater\main.py --hidden-import PyQt5.sip --noconsole
import sys
from PyQt5.QtWidgets import QApplication, QLabel, QWidget
#function imports
from main_page import main_page
from component import grid
from Material_shortage_list import Material_shortage_page, BOM_import_page
from barcode_generator import operation_search_page, out_stock_select_page
from longin_page import log_in_page
from part_code_searcher import frame_search_page
from part_code_generator import add_selection_page


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
window.setWindowTitle("EPR輔助軟體 Version 1.0.6(BETA)")
# window.setFixedWidth(1000)
#place window in (x,y) coordinates
# window.move(2700, 200)
window.setStyleSheet("background: #faffff;")
window.resize(1350, 800)

#display frame 1
# barcode_select_page()
# main_page()
# BOM_import_page()
# frame_search_page()
# Material_shortage_page()
log_in_page()
# frame_search_page()
# add_selection_page()
# out_stock_select_page()

window.setLayout(grid)
grid.setSpacing(10)

window.show()
sys.exit(app.exec()) #terminate the app
