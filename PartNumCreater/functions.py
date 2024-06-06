from PyQt5.QtWidgets import QGridLayout, QLabel, QPushButton, QLineEdit, QComboBox
from PyQt5.QtGui import QPixmap, QCursor
from PyQt5 import QtCore
from urllib.request import urlopen
import json
import pandas as pd
import random


#global dictionary of dynamically changing widgets
widgets = {
    "line_bar1": [],
    "button_input":[],
    "selected_box1":[],
}

#initialliza grid layout
grid = QGridLayout()

def clear_widgets():
    ''' hide all existing widgets and erase
        them from the global dictionary'''
    for widget in widgets:
        if widgets[widget] != []:
            widgets[widget][-1].hide()
        for i in range(0, len(widgets[widget])):
            widgets[widget].pop()

def create_label(name, l_margin, r_margin):
    #create identical buttons with custom left & right margins
    label = QLabel(name)
    label.setFixedWidth(200)
    label.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        '''
            font-size: 20px;
            color: 'black';
        }
        '''
    )
    return label

def create_lineedit(l_margin, r_margin):
    lineedit = QLineEdit()
    lineedit.setFixedWidth(400)
    lineedit.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        '''
            background: '#D9D9D9';
            font-size: 20px;
            color: 'black';
            padding: 5px 0;
            border-radius: 5px;
        }
        *:hover{
            background: '#ECECEC';
        }
        '''
    )
    return lineedit

def create_button(name, color ,l_margin, r_margin):
    button = QPushButton(name)
    button.setFixedWidth(200)
    button.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
    button.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        "background: " + color +";"+
        '''
            font-size: 20px;
            color: 'white';
            padding: 5px 0px;
            border-radius: 5px;
        }
        '''
    )
    # button.clicked.connect(start_game)
    return button
    


#*********************************************
#                  FRAME 1
#*********************************************

def frame1():
    clear_widgets()
    #info widget
    label = create_label("輸入料件資訊", 0, 0)
    grid.addWidget(label, 0, 0, 1, 1)

    #LineEdit widget
    lineEdit = create_lineedit(0,0)
    widgets["line_bar1"].append(lineEdit)
    grid.addWidget(widgets["line_bar1"][-1], 0, 1, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#006262", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 0, 3, 1, 1)

    #info widget
    label = create_label("項目", 0, 0)
    grid.addWidget(label, 1, 0, 1, 1)

    #button widget
    Combox1 = QComboBox()
    Combox1.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
    Combox1.setStyleSheet(
        '''
        *{
            background: '#D9D9D9';
            font-size: 20px;
            color: 'white';
            padding: 5px 0px;
            border-radius: 5px;
        }
        *:hover{
            background: '#ECECEC';
        }
        '''
    )
    #button callback
    widgets["selected_box1"].append(Combox1)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 1, 2, 1, 1)
