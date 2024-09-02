from PyQt5.QtWidgets import QGridLayout
from PyQt5.QtWidgets import QLabel, QPushButton
from PyQt5.QtWidgets import QSizePolicy
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCursor


main_grid = QGridLayout()

main_widget = {
    "nav-bar":[],
    "main-label":[],
    "ost-logo-label":[],
    "ERP-code-button":[],
    "search-button":[],
    "stock-button":[],
    "MRP-button":[],
    "ERP-code-generate-button":[],
    "rule-generate-button":[],
    "diagonally-button":[],
    "ERP-code-layout":[],
}

def clear_widgets(widgets):
    ''' hide all existing widgets and erase
        them from the global dictionary'''
    for widget in widgets:
        if widgets[widget] != []:
            widgets[widget][-1].hide()
        for i in range(0, len(widgets[widget])):
            widgets[widget].pop()

def create_label(name, l_margin, r_margin, font_color, font_size, background_color, align = 'left', radius = 0):
    #create identical buttons with custom left & right margins
    label = QLabel(name)
    # label.setFixedWidth(width)
    # label.setFixedHeight(height)
    label.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        "font-size: " + str(font_size) +"px;"+
        "color: " + font_color+";"+
        "border-radius: " + str(radius) + "px;"+
        "background: " + background_color + ";"
        '''
            padding: 0px 0px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        '''
    )
    if(align == 'center'):
        label.setAlignment(Qt.AlignVCenter | Qt.AlignHCenter)
    elif(align == 'right'):
        label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
    return label

def create_button(name, l_margin, r_margin, font_color, font_size, background_color, hover_font_color, hover_font_size, hover_background_color, radius = 0, hover_pic=None):
    button = QPushButton(name)
    # button.setFixedWidth(width)
    # button.setFixedHeight(height)
    button.setCursor(QCursor(Qt.PointingHandCursor))
    stylesheet = (
        "QPushButton{margin-left: " + str(l_margin) + "px;" +
        "margin-right: " + str(r_margin) + "px;" +
        "background: " + background_color + ";" +
        "font-size: " + str(font_size) + "px;" +
        "color: " + font_color + ";" +
        "border-radius: " + str(radius) + "px;" +
        '''
            padding: 0px;
            font-family: 'Microsoft JhengHei';
            font-weight: bold;
        }
        QPushButton:hover{
        ''' +
        "background: " + hover_background_color + ";" +
        "font-size: " + str(hover_font_size) + "px;" +
        "color: " + hover_font_color + ";"
    )
    
    # 如果 hover_pic 不為 None，則添加 icon 行
    if hover_pic is not None:
        stylesheet += "icon: url(" + hover_pic + ");"
    
    # 關閉 hover 狀態
    stylesheet += "}"
    
    # 設置樣式表
    button.setStyleSheet(stylesheet)

    button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
    return button