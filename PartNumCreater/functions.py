from PyQt5.QtWidgets import QGridLayout, QLabel, QPushButton, QLineEdit, QComboBox, QHBoxLayout
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
    "part_resistance": [],
    "part_kind":[],
    "part_name":[],
}

#save parts all information
itemslist= {}

#global dictionary of dynamically changing widgets
widgets = {
    "label1":[],
    "label2":[],
    "label3":[],
    "label4":[],
    "label5":[],
    "label6":[],
    "label7":[],
    "label8":[],
    "label9":[],
    "label10":[],
    "label11":[],
    "label12":[],
    "label13":[],
    "line_bar1": [],
    "line_bar2": [],
    "line_bar3": [],
    "button_input":[],
    "button_output":[],
    "button_export":[],
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

def clear_parameters(parameters):
    for key in parameters:
        parameters[key].clear()

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
        print(index1, index2)
        frame3(index1, index2)
    elif (index2 == 1):
        print(index1, index2)
        frame4(index1, index2)
    else:
        frame5(index1, index2)

def combo_part_kind_change(index1: int, index2: int, index3: int):
    frame6(index1, index2, index3)


#*********************************************
#                  FRAME 1
#*********************************************

def frame1():
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    #import combo box data
    parameters["part_list"].append(["SMT", "DIP"])

    #info widget
    label1 = create_label("輸入料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 0, 0, 1, 1)

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
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 1, 0, 1, 1)

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
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    #import combo box data
    parameters["part_list"].append(["SMT", "DIP"])
    parameters["part_type"].append(["電容", "電阻", "IC", "橋堆"])

    #info widget
    label1 = create_label("輸入料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 0, 0, 1, 1)

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
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 1, 0, 1, 1)

    #button widget
    Combox1 = create_combobox(parameters["part_list"][-1], 0, 0)
    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_list_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 1, 2, 1, 1)

    #info widget
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 2, 0, 1, 1)

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
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    #import combo box data
    parameters["part_list"].append(["SMT", "DIP"])
    parameters["part_type"].append(["電容", "電阻", "IC", "橋堆"])
    parameters["part_size"].append(["0201","0402"])
    parameters["part_coefficient"].append(["X7R","X5R", "NPO"])
    parameters["part_percentage"].append(["B ± 0.10pF", "C ± 0.25pF"])
    parameters["part_capacity"].append(["102 1000p", "200 20p"])
    parameters["part_voltage"].append(["16V", "25V"])
    parameters["part_manufacturer"].append(["國巨", "華新科"])
    parameters["part_supplier"].append([])
    
    #info widget
    label1 = create_label("輸入料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 0, 0, 1, 1)

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
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 1, 0, 1, 1)

    Combox1 = create_combobox(parameters["part_list"][-1], 0, 0)

    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_list_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 1, 2, 1, 1)

    #2nd
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 2, 0, 1, 1)

    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 2, 2, 1, 1)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].setCurrentIndex(type_choose)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex()))

    #3rd
    label4 = create_label("零件尺寸", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 3, 0, 1, 1)

    Combox3 = create_combobox(parameters["part_size"][-1], 0, 0)
    widgets["selected_box3"].append(Combox3)


    #place global widgets on the grid
    grid.addWidget(widgets["selected_box3"][-1], 3, 2, 1, 1)

    #4th
    label5 = create_label("類型", 0, 0)
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 4, 0, 1, 1)

    Combox4 = create_combobox(parameters["part_coefficient"][-1], 0, 0)
    widgets["selected_box4"].append(Combox4)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box4"][-1], 4, 2, 1, 1)

    #4th
    label6 = create_label("%數", 0, 0)
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 5, 0, 1, 1)

    Combox5 = create_combobox(parameters["part_percentage"][-1], 0, 0)
    widgets["selected_box5"].append(Combox5)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box5"][-1], 5, 2, 1, 1)

    #5th 6th 7th
    label7 = create_label("電容值", 0, 0)
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 6, 0, 1, 1)

    Combox6 = create_combobox(parameters["part_capacity"][-1], 0, 0)
    widgets["selected_box6"].append(Combox6)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box6"][-1], 6, 2, 1, 1)

    #8th
    label8 = create_label("電壓值", 0, 0)
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 7, 0, 1, 1)

    Combox7 = create_combobox(parameters["part_voltage"][-1], 0, 0)
    widgets["selected_box7"].append(Combox7)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box7"][-1], 7, 2, 1, 1)

    #9th
    label9 = create_label("廠商", 0, 0)
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 8, 0, 1, 1)

    Combox8 = create_combobox(parameters["part_manufacturer"][-1], 0, 0)
    widgets["selected_box8"].append(Combox8)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box8"][-1], 8, 2, 1, 1)

    #10th 11th
    label10 = create_label("供應商", 0, 0)
    widgets["label10"].append(label10)
    grid.addWidget(widgets["label10"][-1], 9, 0, 1, 1)

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

    label11 = create_label("品項名稱", 0, 0)
    widgets["label11"].append(label11)
    grid.addWidget(widgets["label11"][-1], 11, 0, 1, 1)

    #LineEdit widget
    lineEdit3 = create_lineedit(0,0)
    widgets["line_bar3"].append(lineEdit3)
    grid.addWidget(widgets["line_bar3"][-1], 11, 1, 1, 2)

    label12 = create_label("已產生編號", 0, 0)
    widgets["label12"].append(label12)
    grid.addWidget(widgets["label12"][-1], 11, 0, 1, 1)

    #QBoxLayout for produced part
    label13 = create_label("", 0, 0)
    widgets["label13"].append(label13)
    grid.addWidget(widgets["label13"][-1], 12, 0, 5, 1)

    hbox_layout = QHBoxLayout()
    hbox_layout.addWidget(QLabel('HBox Label 1'))
    grid.addLayout(hbox_layout, 12, 0, 5, 3)

    button3 = create_button("匯出至Excel", "#1F7145", 0, 0)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 18, 2, 1, 1)

#*********************************************
#                  FRAME 4
#*********************************************

def frame4(part_choose, type_choose):
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    print(parameters)

    #import combo box data
    parameters["part_list"].append(["SMT", "DIP"])
    parameters["part_type"].append(["電容", "電阻", "IC", "橋堆"])
    parameters["part_size"].append(["0201","0402"])
    parameters["part_percentage"].append(["0.50%", "F±1%"])
    parameters["part_resistance"].append(["1KΩ", "560Ω"])
    parameters["part_manufacturer"].append(["國巨", "華新科"])
    parameters["part_supplier"].append([])

    print(parameters)

    #info widget
    label1 = create_label("輸入料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 0, 0, 1, 1)

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
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 1, 0, 1, 1)

    Combox1 = create_combobox(parameters["part_list"][-1], 0, 0)

    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_list_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 1, 2, 1, 1)

    #2nd
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 2, 0, 1, 1)

    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 2, 2, 1, 1)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].setCurrentIndex(type_choose)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex()))

    #3rd
    label4 = create_label("零件尺寸", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 3, 0, 1, 1)

    Combox3 = create_combobox(parameters["part_size"][-1], 0, 0)
    widgets["selected_box3"].append(Combox3)


    #place global widgets on the grid
    grid.addWidget(widgets["selected_box3"][-1], 3, 2, 1, 1)

    #4th
    label5 = create_label("%數", 0, 0)
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 4, 0, 1, 1)

    Combox4 = create_combobox(parameters["part_percentage"][-1], 0, 0)
    widgets["selected_box4"].append(Combox4)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box4"][-1], 4, 2, 1, 1)

    #5th 6th 7th 8th
    label6 = create_label("電阻值", 0, 0)
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 5, 0, 1, 1)

    Combox5 = create_combobox(parameters["part_resistance"][-1], 0, 0)
    widgets["selected_box5"].append(Combox5)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box5"][-1], 5, 2, 1, 1)

    #9th
    label7 = create_label("廠商", 0, 0)
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 6, 0, 1, 1)

    Combox6 = create_combobox(parameters["part_manufacturer"][-1], 0, 0)
    widgets["selected_box6"].append(Combox6)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box6"][-1], 6, 2, 1, 1)

    #10th 11th
    label8 = create_label("供應商", 0, 0)
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 7, 0, 1, 1)

    Combox7 = create_combobox(parameters["part_supplier"][-1], 0, 0)
    widgets["selected_box7"].append(Combox7)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box7"][-1], 7, 2, 1, 1)

    #button widget
    button2 = create_button("生成料號", "#FFFDD4", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_output"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button_output"][-1], 8, 0, 1, 1)

    #LineEdit widget
    lineEdit2 = create_lineedit(0,0)
    widgets["line_bar2"].append(lineEdit2)
    grid.addWidget(widgets["line_bar2"][-1], 8, 1, 1, 2)

    label9 = create_label("品項名稱", 0, 0)
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 11, 0, 1, 1)

    #LineEdit widget
    lineEdit3 = create_lineedit(0,0)
    widgets["line_bar3"].append(lineEdit3)
    grid.addWidget(widgets["line_bar3"][-1], 9, 1, 1, 2)

    label10 = create_label("已產生編號", 0, 0)
    widgets["label10"].append(label10)
    grid.addWidget(widgets["label10"][-1], 9, 0, 1, 1)

    #QBoxLayout for produced part
    label11 = create_label("", 0, 0)
    widgets["label11"].append(label11)
    grid.addWidget(widgets["label11"][-1], 10, 0, 5, 1)

    hbox_layout = QHBoxLayout()
    hbox_layout.addWidget(QLabel('HBox Label 1'))
    grid.addLayout(hbox_layout, 10, 0, 5, 3)

    button3 = create_button("匯出至Excel", "#1F7145", 0, 0)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 16, 2, 1, 1)

#*********************************************
#                  FRAME 5
#*********************************************

def frame5(part_choose, type_choose):
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    print(parameters)

    #import combo box data
    parameters["part_list"].append(["SMT", "DIP"])
    parameters["part_type"].append(["電容", "電阻", "IC", "橋堆"])
    parameters["part_kind"].append(["貼片IC"])

    #info widget
    label1 = create_label("輸入料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 0, 0, 1, 1)

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
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 1, 0, 1, 1)

    #button widget
    Combox1 = create_combobox(parameters["part_list"][-1], 0, 0)
    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_list_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 1, 2, 1, 1)

    #info widget
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 2, 0, 1, 1)

    #button widget
    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].setCurrentIndex(type_choose)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 2, 2, 1, 1)

    #info widget
    label4 = create_label("種類", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 3, 0, 1, 1)
    #button widget
    Combox3 = create_combobox(parameters["part_kind"][-1], 0, 0)
    widgets["selected_box3"].append(Combox3)
    widgets["selected_box3"][-1].currentIndexChanged.connect(lambda: combo_part_kind_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box3"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box3"][-1], 3, 2, 1, 1)

def frame6(part_choose, type_choose, kind_choose):
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    print(parameters)

    #import combo box data
    parameters["part_list"].append(["SMT", "DIP"])
    parameters["part_type"].append(["電容", "電阻", "IC", "橋堆"])
    parameters["part_kind"].append(["貼片IC"])
    parameters["part_name"].append(["RS621KXF", "LTL431ALT1G(乐山)", "ID5S609SEC-R1", "ULN2003G(UTC)"])
    parameters["part_manufacturer"].append([])
    parameters["part_supplier"].append([])

    #info widget
    label1 = create_label("輸入料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 0, 0, 1, 1)

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
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 1, 0, 1, 1)

    #button widget
    Combox1 = create_combobox(parameters["part_list"][-1], 0, 0)
    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_list_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 1, 2, 1, 1)

    #info widget
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 2, 0, 1, 1)

    #button widget
    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].setCurrentIndex(type_choose)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 2, 2, 1, 1)

    #info widget
    label4 = create_label("種類", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 3, 0, 1, 1)
    #button widget
    Combox3 = create_combobox(parameters["part_kind"][-1], 0, 0)
    widgets["selected_box3"].append(Combox3)
    widgets["selected_box3"][-1].setCurrentIndex(kind_choose)
    widgets["selected_box3"][-1].currentIndexChanged.connect(lambda: combo_part_kind_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box3"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box3"][-1], 3, 2, 1, 1)

    #info widget
    label5 = create_label("料件名稱", 0, 0)
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 4, 0, 1, 1)

    Combox4 = create_combobox(parameters["part_name"][-1], 0, 0)
    widgets["selected_box4"].append(Combox4)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box4"][-1], 4, 2, 1, 1)

    #9th
    label6 = create_label("廠商", 0, 0)
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 5, 0, 1, 1)

    Combox5 = create_combobox(parameters["part_manufacturer"][-1], 0, 0)
    widgets["selected_box5"].append(Combox5)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box5"][-1], 5, 2, 1, 1)

    #10th 11th
    label7 = create_label("供應商", 0, 0)
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 6, 0, 1, 1)

    Combox6 = create_combobox(parameters["part_supplier"][-1], 0, 0)
    widgets["selected_box6"].append(Combox6)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box6"][-1], 6, 2, 1, 1)

    #button widget
    button2 = create_button("生成料號", "#FFFDD4", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_output"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button_output"][-1], 7, 0, 1, 1)

    #LineEdit widget
    lineEdit2 = create_lineedit(0,0)
    widgets["line_bar2"].append(lineEdit2)
    grid.addWidget(widgets["line_bar2"][-1], 7, 1, 1, 2)

    label8 = create_label("品項名稱", 0, 0)
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 11, 0, 1, 1)

    #LineEdit widget
    lineEdit3 = create_lineedit(0,0)
    widgets["line_bar3"].append(lineEdit3)
    grid.addWidget(widgets["line_bar3"][-1], 8, 1, 1, 2)

    label9 = create_label("已產生編號", 0, 0)
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 8, 0, 1, 1)

    #QBoxLayout for produced part
    label10 = create_label("", 0, 0)
    widgets["label10"].append(label10)
    grid.addWidget(widgets["label10"][-1], 9, 0, 5, 1)

    hbox_layout = QHBoxLayout()
    hbox_layout.addWidget(QLabel('HBox Label 1'))
    grid.addLayout(hbox_layout, 9, 0, 5, 3)

    button3 = create_button("匯出至Excel", "#1F7145", 0, 0)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 15, 2, 1, 1)



    
    