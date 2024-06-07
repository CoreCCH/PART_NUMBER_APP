from PyQt5.QtWidgets import QGridLayout, QLabel, QPushButton, QLineEdit, QComboBox
from PyQt5.QtGui import QPixmap, QCursor
from PyQt5 import QtCore
from urllib.request import urlopen
import json
import pandas as pd
import random

#dictionary to store local pre-load parameters on a global level
parameters = {
    "part_list": [],
    "part_type": [],
    "part_size": [],
    "part_coefficient": [],
    "part_percentage": [],
    "part_capacity": [],
    "part_voltage": [],
    "part_manufacturer": [],
    "part_supplier": [],
}

#save parts all information
itemslist= {}

#global dictionary of dynamically changing widgets
widgets = {
    "label1":[],
    "label2":[],
    "label3":[],
    "line_bar1": [],
    "line_bar2": [],
    "line_bar3": [],
    "button_input":[],
    "button_output":[],
    "selected_box1":[],
    "selected_box2":[],
    "selected_box3":[],
    "selected_box4":[],
    "selected_box5":[],
    "selected_box6":[],
    "selected_box7":[],
    "selected_box8":[],
    "selected_box9":[],
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
            color: 'black';
            padding: 5px 0px;
            border-radius: 5px;
        }
        '''
    )
    # button.clicked.connect(start_game)
    return button
    
def create_combobox(item_list, l_margin, r_margin):
    combo = QComboBox()
    combo.setFixedWidth(200)
    combo.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
    combo.addItems(item_list)
    combo.setCurrentIndex(-1)
    combo.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        '''
            background: '#D9D9D9';
            font-size: 20px;
            color: 'black';
            padding: 5px 0px;
            border-radius: 5px;
        }
        *:hover{
            background: '#ECECEC';
        }
        '''
    )
    return combo

def combo_part_list_change(index: int):
    frame2(index)

def combo_part_type_change(index1: int, index2: int):
    if (index2 == 0):
        frame3(index1, index2)


#*********************************************
#                  FRAME 1
#*********************************************

def frame1():
    clear_widgets()

    global parameters

    #import combo box data
    parameters["part_list"].append(["SMT", "DIP"])

    #info widget
    label = create_label("輸入料件資訊", 0, 0)
    grid.addWidget(label, 0, 0, 1, 1)

    #LineEdit widget
    lineEdit = create_lineedit(0,0)
    widgets["line_bar1"].append(lineEdit)
    grid.addWidget(widgets["line_bar1"][-1], 0, 1, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#01C7C7", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 0, 3, 1, 1)

    #info widget
    label = create_label("項目", 0, 0)
    grid.addWidget(label, 1, 0, 1, 1)

    #button widget
    Combox1 = create_combobox(parameters["part_list"][-1], 0, 0)
    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_list_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 1, 2, 1, 1)


#*********************************************
#                  FRAME 2
#*********************************************

def frame2(part_choose):
    clear_widgets()

    global parameters

    #import combo box data
    parameters["part_type"].append(["電容", "電阻", "IC", "橋堆"])

    #info widget
    label = create_label("輸入料件資訊", 0, 0)
    grid.addWidget(label, 0, 0, 1, 1)

    #LineEdit widget
    lineEdit = create_lineedit(0,0)
    widgets["line_bar1"].append(lineEdit)
    grid.addWidget(widgets["line_bar1"][-1], 0, 1, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#01C7C7", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 0, 3, 1, 1)

    #info widget
    label = create_label("項目", 0, 0)
    grid.addWidget(label, 1, 0, 1, 1)

    #button widget
    Combox1 = create_combobox(parameters["part_list"][-1], 0, 0)
    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_list_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 1, 2, 1, 1)

    #info widget
    label = create_label("種類", 0, 0)
    grid.addWidget(label, 2, 0, 1, 1)

    #button widget
    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 2, 2, 1, 1)


#*********************************************
#                  FRAME 3
#*********************************************

def frame3(part_choose, type_choose):
    clear_widgets()

    global parameters

    #import combo box data
    parameters["part_size"].append(["0201","0402"])
    parameters["part_coefficient"].append(["X7R","X5R", "NPO"])
    parameters["part_percentage"].append(["B ± 0.10pF", "C ± 0.25pF"])
    parameters["part_capacity"].append(["102 1000p", "200 20p"])
    parameters["part_voltage"].append(["16V", "25V"])
    parameters["part_manufacturer"].append(["國巨", "華新科"])
    parameters["part_supplier"].append([])
    
    #info widget
    label = create_label("輸入料件資訊", 0, 0)
    grid.addWidget(label, 0, 0, 1, 1)

    #LineEdit widget
    lineEdit1 = create_lineedit(0,0)
    widgets["line_bar1"].append(lineEdit1)
    grid.addWidget(widgets["line_bar1"][-1], 0, 1, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#01C7C7", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 0, 3, 1, 1)

    #1st
    label = create_label("項目", 0, 0)
    grid.addWidget(label, 1, 0, 1, 1)

    Combox1 = create_combobox(parameters["part_list"][-1], 0, 0)

    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_list_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 1, 2, 1, 1)

    #2nd
    label = create_label("種類", 0, 0)
    grid.addWidget(label, 2, 0, 1, 1)

    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 2, 2, 1, 1)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].setCurrentIndex(type_choose)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box2"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex()))

    #3rd
    label = create_label("零件尺寸", 0, 0)
    grid.addWidget(label, 3, 0, 1, 1)

    Combox3 = create_combobox(parameters["part_size"][-1], 0, 0)
    widgets["selected_box3"].append(Combox3)


    #place global widgets on the grid
    grid.addWidget(widgets["selected_box3"][-1], 3, 2, 1, 1)

    #4th
    label = create_label("類型", 0, 0)
    grid.addWidget(label, 4, 0, 1, 1)

    Combox4 = create_combobox(parameters["part_coefficient"][-1], 0, 0)
    widgets["selected_box4"].append(Combox4)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box4"][-1], 4, 2, 1, 1)

    #4th
    label = create_label("%數", 0, 0)
    grid.addWidget(label, 5, 0, 1, 1)

    Combox5 = create_combobox(parameters["part_percentage"][-1], 0, 0)
    widgets["selected_box5"].append(Combox5)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box5"][-1], 5, 2, 1, 1)

    #5th 6th 7th
    label = create_label("電容值", 0, 0)
    grid.addWidget(label, 6, 0, 1, 1)

    Combox6 = create_combobox(parameters["part_capacity"][-1], 0, 0)
    widgets["selected_box6"].append(Combox6)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box6"][-1], 6, 2, 1, 1)

    #8th
    label = create_label("電壓值", 0, 0)
    grid.addWidget(label, 7, 0, 1, 1)

    Combox7 = create_combobox(parameters["part_voltage"][-1], 0, 0)
    widgets["selected_box7"].append(Combox7)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box7"][-1], 7, 2, 1, 1)

    #9th
    label = create_label("廠商", 0, 0)
    grid.addWidget(label, 8, 0, 1, 1)

    Combox8 = create_combobox(parameters["part_manufacturer"][-1], 0, 0)
    widgets["selected_box8"].append(Combox8)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box8"][-1], 8, 2, 1, 1)

    #10th 11th
    label = create_label("供應商", 0, 0)
    grid.addWidget(label, 9, 0, 1, 1)

    Combox9 = create_combobox(parameters["part_supplier"][-1], 0, 0)
    widgets["selected_box9"].append(Combox9)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box9"][-1], 9, 2, 1, 1)

    #button widget
    button2 = create_button("生成料號", "#FFFDD4", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_output"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button_output"][-1], 10, 0, 1, 1)

    #LineEdit widget
    lineEdit2 = create_lineedit(0,0)
    widgets["line_bar2"].append(lineEdit2)
    grid.addWidget(widgets["line_bar2"][-1], 10, 1, 1, 2)

    label = create_label("品項名稱", 0, 0)
    grid.addWidget(label, 11, 0, 1, 1)

    #LineEdit widget
    lineEdit3 = create_lineedit(0,0)
    widgets["line_bar3"].append(lineEdit3)
    grid.addWidget(widgets["line_bar3"][-1], 11, 1, 1, 2)

    label = create_label("已產生編號", 0, 0)
    grid.addWidget(label, 12, 0, 1, 1)

    
    