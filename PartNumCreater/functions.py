from PyQt5.QtWidgets import QGridLayout, QLabel, QPushButton, QLineEdit, QComboBox, QScrollArea, QDialog, QVBoxLayout, QGraphicsDropShadowEffect, QTextEdit
from PyQt5.QtGui import QPixmap, QCursor, QColor
from PyQt5 import QtCore
from PyQt5.QtCore import Qt
from urllib.request import urlopen
import pandas as pd
import execl_handle
from datetime import datetime

df = execl_handle.excel_file_read('曜璿東命名規則 20240605-2.xlsx', '命名規則')
df2 = execl_handle.excel_file_read('曜璿東命名規則 20240605-2.xlsx', '電容種類規則')

#import combo box data
object_data = execl_handle.excel_list_read(df)
object_selected = [list(d.keys())[0] for d in object_data]
object_selection, type_data = execl_handle.excel_type_read(df)
size_data = execl_handle.execl_size_read(df, object_selection)
percentage_data = execl_handle.execl_percentage_read(df, object_selection)
capacity_data = execl_handle.execl_capacity_read(df, object_selection)
voltage_data = execl_handle.execl_voltage_read(df, object_selection)
resistance_data = execl_handle.execl_capacity_read(df, object_selection)
manufacturer_data = execl_handle.execl_manufacturer_read(df, object_selection)
supplier_data = execl_handle.execl_supplier_read(df, object_selection)
kind_data = execl_handle.execl_size_read(df, object_selection)
name_data = execl_handle.execl_capacity_read(df, object_selection)
capacity_change = execl_handle.add_capacity(df2)

headers = ["part_number", "品項編號", "品項名稱", "項目", "種類", "尺寸/種類","%數", "容值/阻值/名稱", "電壓", "廠商", "供應商", "產生時間"]
file_path = 'output.xlsx'
output_data = {}
output_df = execl_handle.check_output_existing(file_path, headers)
check_buffer = [output_df.iloc[:, 0][i] for i in range(output_df.iloc[:, 0].size)] if output_df is not None else []
part_code = int([output_df.iloc[:, 1][i] for i in range(output_df.iloc[:, 0].size)][-1].split("C")[-1])+1 if output_df is not None else 1

vbox_label_layout = QVBoxLayout()
# vbox_trash_layout = QVBoxLayout()
Pannel = None


#dictionary to store local pre-load parameters on a global level
parameters = {
    "part_object": [],
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

part_number_label = {}
part_number_button= {}

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
    label.setFixedWidth(150)
    label.setFixedHeight(44)
    label.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        '''
            font-size: 20px;
            color: 'black';
            padding: 0px 0px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        '''
    )
    return label

def create_lineedit(l_margin, r_margin):
    lineedit = QLineEdit()
    lineedit.setFixedHeight(44)
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
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        *:hover{
            background: '#ECECEC';
            font-size: 20px;
            color: 'black';
            padding: 5px 0;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        '''
    )
    return lineedit

def create_button(name, color ,l_margin, r_margin):
    button = QPushButton(name)
    button.setFixedWidth(110)
    button.setFixedHeight(44)
    button.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
    button.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        "background: " + color +";"+
        '''
            font-size: 20px;
            color: 'black';
            padding: 0px 0px;
            border-radius: 10px;
            font-family: 'Microsoft JhengHei';
            font-weight: bold;
        }
        '''
    )
    # button.clicked.connect(start_game)
    shadow_effect = QGraphicsDropShadowEffect()
    shadow_effect.setBlurRadius(10)
    shadow_effect.setXOffset(5)
    shadow_effect.setYOffset(5)
    shadow_effect.setColor(QColor(0, 0, 0, 160))
    button.setGraphicsEffect(shadow_effect)
    return button
    
def create_combobox(item_list, l_margin, r_margin):
    combo = QComboBox()
    combo.setFixedWidth(200)
    combo.setFixedHeight(44)
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
            font-size: 14px;
            font-family: Microsoft JhengHei
        }
        *:hover{
            background: '#ECECEC';
        }
        '''
    )
    return combo

def combo_part_object_change(index: int):
    frame2(index)

def combo_part_type_change(index1: int, index2: int, text1: str, text2: str):
    if (text1 == "SMT" and text2 == "電容"):
        frame3(index1, index2)
    elif (text1 == "SMT" and  text2 == "電阻"):
        frame4(index1, index2)
    else:
        frame5(index1, index2)

def combo_part_kind_change(index1: int, index2: int, index3: int, text1: str, text2: str, text3: str):
    if ("電容" in text3):
        frame7(index1, index2, index3)
    else:    
        frame6(index1, index2, index3)

def part_number_generate_for_frame3(line_edit1, line_edit2, object: str, type: str, size: str, coefficient:str, percentage: str, capacity: str, voltage: str, manufacturer: str, supplier: str):
    # capacity generate
    part_number = ""

    check = True

    if (object == ""):
        show_alert("請選擇料件項目")
        check = False
    if(type == ""):
        show_alert("請選擇料件種類")
        check = False
    if(size == ""):
        show_alert("請選擇料件尺寸")
        check = False
    if(coefficient == ""): 
        show_alert("請選擇料件係數")
        check = False 
    if(percentage == ""):
        show_alert("請選擇料件%數")
        check = False
    if(capacity == ""):
        show_alert("請選擇料件容值")
        check = False
    if(voltage == ""):
        show_alert("請選擇料件電壓")
        check = False
    if(check == False): return False

    part_number += str(next((item[object] for item in object_data if object in item), None))

    part_number += str(next((item[type] for item in type_data[object] if type in item), None))

    part_number += str(next((item[size] for item in size_data[type] if size in item), None))

    part_number += capacity_change[coefficient][capacity_change["X7R"].index(next((item[percentage] for item in percentage_data[type] if percentage in item), None))] if coefficient in capacity_change else ""

    part_number += str(next((item[capacity] for item in capacity_data[type] if capacity in item), None))

    part_number += str(next((item[voltage] for item in voltage_data[type] if voltage in item), None))

    part_number += '0' if manufacturer == '' else str(next((item[manufacturer] for item in manufacturer_data[type] if manufacturer in item), None))
   
    part_number += '00' if supplier == '' else str(next((item[supplier] for item in supplier_data[type] if supplier in item), None))

    line_edit1.setText(part_number)

    part_name= ""

    part_name= coefficient+","+str(next((item[capacity] for item in capacity_data[type] if capacity in item), None))+"/"+voltage+","+percentage.split("±")[-1]+","+size+" ("+ manufacturer +")"

    line_edit2.setText(part_name)

    current_time = datetime.now()

    global output_data
    global part_code
    
    if (part_number not in check_buffer):
        if (part_number not in output_data):
            output_data.update({part_number:["C"+str(part_code).zfill(5), part_name, object, type, size, percentage, voltage, capacity, manufacturer,supplier, str(current_time.date())+" "+f"{current_time.hour}:{current_time.minute}:{current_time.second}"]})
            part_code += 1
            global Pannel
            Pannel.append(part_number)
    else:
        show_alert("已存在料號 ",part_number)

def part_number_generate_for_frame4(line_edit1, line_edit2, object: str, type: str, size: str, percentage: str, resistance: str, manufacturer: str, supplier: str):
    # capacity generate
    part_number = ""

    check = True
    
    if (object == ""):
        show_alert("請選擇料件項目")
        check = False
    if(type == ""):
        show_alert("請選擇料件種類")
        check = False
    if(size == ""):
        show_alert("請選擇料件尺寸")
        check = False
    if(percentage == ""):
        show_alert("請選擇料件%數")
        check = False
    if(resistance == ""):
        show_alert("請選擇料件組值")
        check = False
    if(check == False): return False

    part_number += str(next((item[object] for item in object_data if object in item), None))

    part_number += str(next((item[type] for item in type_data[object] if type in item), None))

    part_number += str(next((item[size] for item in size_data[type] if size in item), None))

    part_number += str(next((item[percentage] for item in percentage_data[type] if percentage in item), None))

    part_number += str(next((item[resistance] for item in resistance_data[type] if resistance in item), None))

    part_number += '0' if manufacturer == '' else str(next((item[manufacturer] for item in manufacturer_data[type] if manufacturer in item), None))
   
    part_number += '00' if supplier == '' else str(next((item[supplier] for item in supplier_data[type] if supplier in item), None))

    line_edit1.setText(part_number)

    part_name= ""

    part_name= resistance.split("Ω")[0]+","+percentage.split("±")[-1]+",1/4W,"+","+size+" ("+ manufacturer +")"

    line_edit2.setText(part_name)

    current_time = datetime.now()

    global output_data
    global part_code
    
    if (part_number not in check_buffer):
        if (part_number not in output_data):
            output_data.update({part_number:["C"+str(part_code).zfill(5), part_name, object, type, size, percentage, "", resistance, manufacturer,supplier, str(current_time.date())+" "+f"{current_time.hour}:{current_time.minute}:{current_time.second}"]})
            part_code += 1
            global Pannel
            Pannel.append(part_number)
    else:
        show_alert("已存在料號 ",part_number)

def part_number_generate_for_frame6(line_edit1, object: str, type: str, kind: str, name: str, manufacturer: str, supplier: str):
    # capacity generate
    part_number = ""

    check = True

    if (object == ""):
        show_alert("請選擇料件項目")
        check = False
    if(type == ""):
        show_alert("請選擇料件種類")
        check = False
    if(kind == ""):
        show_alert("請選擇料件種類")
        check = False
    if(name == ""):
        show_alert("請選擇料件名稱")
        check = False
    if(check == False): return False

    part_number += str(next((item[object] for item in object_data if object in item), None))

    part_number += str(next((item[type] for item in type_data[object] if type in item), None))

    

    part_number += str(next((item[kind] for item in kind_data[type] if kind in item), None))

    part_number += str(next((item[name] for item in name_data[type] if name in item), None)).zfill(5)

    part_number += '0' if manufacturer == '' else str(next((item[manufacturer] for item in manufacturer_data[type] if manufacturer in item), None))
   
    part_number += '00' if supplier == '' else str(next((item[supplier] for item in supplier_data[type] if supplier in item), None))

    line_edit1.setText(part_number)

    current_time = datetime.now()

    global output_data
    global part_code
    global vbox_label_layout
    global scrollArea
    # global vbox_trash_layout
    
    if (part_number not in check_buffer):
        if (part_number not in output_data):
            output_data.update({part_number:["C"+str(part_code).zfill(5), "", object, type, kind, "", name, "", manufacturer,supplier, str(current_time.date())+" "+f"{current_time.hour}:{current_time.minute}:{current_time.second}"]})
            part_code += 1
            # globals()['label'+part_number] = create_label(part_number,0,0)
            # globals()['button'+part_number] = create_button("delete","#DFDFDF",0,0)
            # vbox_label_layout.addWidget(globals()['label'+part_number])
            # part_name_list = ""
            global Pannel
            Pannel.append(part_number)
            
            # vbox_trash_layout.addWidget(globals()['button'+part_number])
    else:
        show_alert("已存在料號 ",part_number)

def part_number_generate_for_frame7(line_edit1, object: str, type: str, kind: str, percentage: str, name: str, voltage: str, manufacturer: str, supplier: str):

    # capacity generate
    part_number = ""
    
    check = True

    if (object == ""):
        show_alert("請選擇料件項目")
        check = False
    if(type == ""):
        show_alert("請選擇料件種類")
        check = False
    if(kind == ""):
        show_alert("請選擇料件種類")
        check = False
    if(name == ""):
        show_alert("請選擇料件名稱")
        check = False 
    if(voltage == ""):
        show_alert("請選擇料件電壓")
        check = False
    if(check == False): return False
  
    part_number += str(next((item[object] for item in object_data if object in item), None))

    part_number += str(next((item[type] for item in type_data[object] if type in item), None))

    part_number += str(next((item[kind] for item in kind_data[type] if kind in item), None))

    part_number += '0' if percentage == '' else str(next((item[percentage] for item in percentage_data[type] if percentage in item), None))

    part_number += str(next((item[name] for item in name_data[type] if name in item), None))

    part_number += str(next((item[voltage] for item in voltage_data[type] if voltage in item), None))

    part_number += '0' if manufacturer == '' else str(next((item[manufacturer] for item in manufacturer_data[type] if manufacturer in item), None))

    part_number += '00' if supplier == '' else str(next((item[supplier] for item in supplier_data[type] if supplier in item), None))

    line_edit1.setText(part_number)

    current_time = datetime.now()

    global output_data
    global part_code
    
    if (part_number not in check_buffer):
        if (part_number not in output_data):
            output_data.update({part_number:["C"+str(part_code).zfill(5), "", object, type, kind, percentage, name, voltage, manufacturer,supplier, str(current_time.date())+" "+f"{current_time.hour}:{current_time.minute}:{current_time.second}"]})
            part_code += 1
            global Pannel
            Pannel.append(part_number)
    else:
        show_alert("已存在料號 ",part_number)
        


def export_data_to_excel():
    for part_number, values in list(output_data.items()):
        execl_handle.append_data_to_excel(file_path, part_number, values, headers)
        del output_data[part_number]
        global Pannel
    
    Pannel.clear()
    for part_number in output_data:
        Pannel.append(part_number)
    
    global check_buffer
    output_df = execl_handle.check_output_existing(file_path, headers)
    check_buffer = [output_df.iloc[:, 0][i] for i in range(output_df.iloc[:, 0].size)] if output_df is not None else []

def show_alert(message):
    # 创建对话框
    dialog = QDialog()
    dialog.setWindowTitle("警告")
    dialog.resize(200,100)
    layout = QVBoxLayout()
    
    # 为每个已存在的料号添加一个标签
    layout.addWidget(QLabel(message))
    
    # 添加关闭按钮
    close_button = QPushButton("Close")
    close_button.clicked.connect(dialog.close)
    layout.addWidget(close_button)
    
    # 设置对话框布局
    dialog.setLayout(layout)
    
    # 显示对话框并运行应用程序事件循环
    dialog.exec_()
    
#*********************************************
#                  FRAME 1
#*********************************************

def frame1():
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    #import combo box data
    data = execl_handle.excel_list_read(df)
    parameters["part_object"].append([list(d.keys())[0] for d in data])

    #info widget
    label1 = create_label("料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 1, 1, 1, 1)

    #LineEdit widget
    lineEdit = create_lineedit(0,0)
    
    widgets["line_bar1"].append(lineEdit)
    grid.addWidget(widgets["line_bar1"][-1], 1, 2, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#008E8E", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 1, 4, 1, 1)

    #info widget
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 2, 1, 1, 1)

    #button widget
    Combox1 = create_combobox(parameters["part_object"][-1], 0, 0)
    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_object_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 2, 2, 1, 1)

    label11 = create_label("已產生編號", 0, 0)
    widgets["label11"].append(label11)
    grid.addWidget(widgets["label11"][-1], 2, 6, 1, 1)

    #QBoxLayout for produced part
    # label12 = create_label("", 0, 0)
    # label12.setStyleSheet("background: #DFDFDF")
    # widgets["label12"].append(label12)
    # grid.addWidget(widgets["label12"][-1], 3, 6, 4, 1)
    global Pannel
    Pannel = QTextEdit()
    Pannel.setReadOnly(True)
    Pannel.setStyleSheet("font-size:20px; font-family: Microsoft JhengHei;font-weight: bold;")
    Pannel.setFixedWidth(180)
    grid.addWidget(Pannel, 3, 6, 4, 1)
    # grid.addLayout(vbox_trash_layout, 3, 7, 4, 1)

    

    button3 = create_button("匯出至Excel", "#1F7145", 0, 0)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 7, 6, 1, 1)


#*********************************************
#                  FRAME 2
#*********************************************

def frame2(part_choose):
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    #import combo box data
    parameters["part_object"].append(object_selected)

    parameters["part_type"].append([list(d.keys())[0] for d in type_data[object_selected[part_choose]]])

    #info widget
    label1 = create_label("料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 1, 1, 1, 1)

    #LineEdit widget
    lineEdit = create_lineedit(0,0)
    widgets["line_bar1"].append(lineEdit)
    grid.addWidget(widgets["line_bar1"][-1], 1, 2, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#008E8E", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 1, 4, 1, 1)

    #info widget
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 2, 1, 1, 1)

    #button widget
    Combox1 = create_combobox(parameters["part_object"][-1], 0, 0)
    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_object_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 2, 2, 1, 1)

    #info widget
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 3, 1, 1, 1)

    #button widget
    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 3, 2, 1, 1)

    label11 = create_label("已產生編號", 0, 0)
    widgets["label11"].append(label11)
    grid.addWidget(widgets["label11"][-1], 2, 6, 1, 1)

    #QBoxLayout for produced part
    # label12 = create_label("", 0, 0)
    # label12.setStyleSheet("background: #DFDFDF")
    # widgets["label12"].append(label12)
    # grid.addWidget(widgets["label12"][-1], 3, 6, 4, 1)

    grid.addWidget(Pannel, 3, 6, 4, 1)
    # grid.addLayout(vbox_trash_layout, 3, 6, 4, 1)

    button3 = create_button("匯出至Excel", "#1F7145", 0, 0)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 7, 6, 1, 1)

#*********************************************
#                  FRAME 3
#*********************************************

def frame3(part_choose, type_choose):
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    
    parameters["part_object"].append(object_selected)

    parameters["part_type"].append([list(d.keys())[0] for d in type_data[object_selected[part_choose]]])
    
    #info widget
    label1 = create_label("料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 1, 1, 1, 1)

    #LineEdit widget
    lineEdit1 = create_lineedit(0,0)
    widgets["line_bar1"].append(lineEdit1)
    grid.addWidget(widgets["line_bar1"][-1], 1, 2, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#008E8E", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 1, 4, 1, 1)

    #1st
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 2, 1, 1, 1)

    Combox1 = create_combobox(parameters["part_object"][-1], 0, 0)

    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_object_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 2, 2, 1, 1)

    #2nd
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 3, 1, 1, 1)

    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 3, 2, 1, 1)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].setCurrentIndex(type_choose)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText()))

    try:
        parameters["part_size"].append([list(d.keys())[0] for d in size_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_size"].append(["0201","0402"])

    parameters["part_coefficient"].append(["X7R","X5R", "NPO"])

    try:
        parameters["part_percentage"].append([list(d.keys())[0] for d in percentage_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_percentage"].append(["B ± 0.10pF", "C ± 0.25pF"])

    try:
        parameters["part_capacity"].append([list(d.keys())[0] for d in capacity_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_capacity"].append(["102 1000p", "200 20p"])

    try:
        parameters["part_voltage"].append([list(d.keys())[0] for d in voltage_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_voltage"].append(["16V", "25V"])

    try:
        parameters["part_manufacturer"].append([list(d.keys())[0] for d in manufacturer_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_manufacturer"].append(["國巨", "華新科"])

    try:
        parameters["part_supplier"].append([list(d.keys())[0] for d in supplier_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_supplier"].append([])

    #3rd
    label4 = create_label("零件尺寸", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 4, 1, 1, 1)

    Combox3 = create_combobox(parameters["part_size"][-1], 0, 0)
    widgets["selected_box3"].append(Combox3)


    #place global widgets on the grid
    grid.addWidget(widgets["selected_box3"][-1], 4, 2, 1, 1)

    #4th
    label5 = create_label("類型", 0, 0)
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 5, 1, 1, 1)

    Combox4 = create_combobox(parameters["part_coefficient"][-1], 0, 0)
    widgets["selected_box4"].append(Combox4)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box4"][-1], 5, 2, 1, 1)

    #4th
    label6 = create_label("%數", 0, 0)
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 6, 1, 1, 1)

    Combox5 = create_combobox(parameters["part_percentage"][-1], 0, 0)
    widgets["selected_box5"].append(Combox5)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box5"][-1], 6, 2, 1, 1)

    #5th 6th 7th
    label7 = create_label("電容值", 0, 0)
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 2, 3, 1, 1)

    Combox6 = create_combobox(parameters["part_capacity"][-1], 0, 0)
    widgets["selected_box6"].append(Combox6)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box6"][-1], 2, 4, 1, 1)

    #8th
    label8 = create_label("電壓值", 0, 0)
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 3, 3, 1, 1)

    Combox7 = create_combobox(parameters["part_voltage"][-1], 0, 0)
    widgets["selected_box7"].append(Combox7)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box7"][-1], 3, 4, 1, 1)

    #9th
    label9 = create_label("廠商", 0, 0)
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 4, 3, 1, 1)

    Combox8 = create_combobox(parameters["part_manufacturer"][-1], 0, 0)
    widgets["selected_box8"].append(Combox8)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box8"][-1], 4, 4, 1, 1)

    #10th 11th
    label10 = create_label("供應商", 0, 0)
    widgets["label10"].append(label10)
    grid.addWidget(widgets["label10"][-1], 5, 3, 1, 1)

    Combox9 = create_combobox(parameters["part_supplier"][-1], 0, 0)
    widgets["selected_box9"].append(Combox9)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box9"][-1], 5, 4, 1, 1)

    #button widget
    button2 = create_button("生成料號", "#FFFDD4", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_output"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button_output"][-1], 7, 1, 1, 1)

    #LineEdit widget
    lineEdit2 = create_lineedit(0,0)
    lineEdit2.setReadOnly(True)
    widgets["line_bar2"].append(lineEdit2)
    grid.addWidget(widgets["line_bar2"][-1], 7, 2, 1, 3)

    label11 = create_label("品項名稱", 0, 0)
    widgets["label11"].append(label11)
    grid.addWidget(widgets["label11"][-1], 8, 1, 1, 1)

    #LineEdit widget
    lineEdit3 = create_lineedit(0,0)
    lineEdit3.setReadOnly(True)
    widgets["line_bar3"].append(lineEdit3)
    grid.addWidget(widgets["line_bar3"][-1], 8, 2, 1, 3)

    label12 = create_label("已產生編號", 0, 0)
    widgets["label12"].append(label12)
    grid.addWidget(widgets["label12"][-1], 2, 6, 1, 1)

    #QBoxLayout for produced part
    # label13 = create_label("", 0, 0)
    # label13.setStyleSheet("background: #DFDFDF")
    # widgets["label13"].append(label13)
    # grid.addWidget(widgets["label13"][-1], 3, 6, 4, 1)

    grid.addWidget(Pannel, 3, 6, 4, 1)
    # grid.addLayout(vbox_trash_layout, 3, 6, 4, 1)

    button3 = create_button("匯出至Excel", "#1F7145", 0, 0)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 7, 6, 1, 1)

    button2.clicked.connect(lambda: part_number_generate_for_frame3(lineEdit2, lineEdit3, widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText(), widgets["selected_box3"][-1].currentText(), widgets["selected_box4"][-1].currentText(), widgets["selected_box5"][-1].currentText(), widgets["selected_box6"][-1].currentText(), widgets["selected_box7"][-1].currentText(), widgets["selected_box8"][-1].currentText(), widgets["selected_box9"][-1].currentText()))
    button3.clicked.connect(export_data_to_excel)

#*********************************************
#                  FRAME 4
#*********************************************

def frame4(part_choose, type_choose):
    global parameters

    clear_widgets()

    clear_parameters(parameters)


    #import combo box data
    parameters["part_object"].append(object_selected)

    parameters["part_type"].append([list(d.keys())[0] for d in type_data[object_selected[part_choose]]])

    #info widget
    label1 = create_label("料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 1, 1, 1, 1)

    #LineEdit widget
    lineEdit = create_lineedit(0,0)
    widgets["line_bar1"].append(lineEdit)
    grid.addWidget(widgets["line_bar1"][-1], 1, 2, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#008E8E", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 1, 4, 1, 1)

    #info widget
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 2, 1, 1, 1)

    Combox1 = create_combobox(parameters["part_object"][-1], 0, 0)

    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_object_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 2, 2, 1, 1)

    #2nd
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 3, 1, 1, 1)

    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 3, 2, 1, 1)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].setCurrentIndex(type_choose)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText()))

    
    try:
        parameters["part_size"].append([list(d.keys())[0] for d in size_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_size"].append(["0201","0402"])

    try:
        parameters["part_percentage"].append([list(d.keys())[0] for d in percentage_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_percentage"].append(["0.50%", "F±1%"])

    try:
        parameters["part_resistance"].append([list(d.keys())[0] for d in resistance_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_resistance"].append(["1KΩ", "560Ω"])

    try:
        parameters["part_manufacturer"].append([list(d.keys())[0] for d in manufacturer_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_manufacturer"].append(["國巨", "華新科"])

    try:
        parameters["part_supplier"].append([list(d.keys())[0] for d in supplier_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_supplier"].append([])

    #3rd
    label4 = create_label("零件尺寸", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 4, 1, 1, 1)

    Combox3 = create_combobox(parameters["part_size"][-1], 0, 0)
    widgets["selected_box3"].append(Combox3)


    #place global widgets on the grid
    grid.addWidget(widgets["selected_box3"][-1], 4, 2, 1, 1)

    #4th
    label5 = create_label("%數", 0, 0)
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 5, 1, 1, 1)

    Combox4 = create_combobox(parameters["part_percentage"][-1], 0, 0)
    widgets["selected_box4"].append(Combox4)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box4"][-1], 5, 2, 1, 1)

    #5th 6th 7th 8th
    label6 = create_label("電阻值", 0, 0)
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 2, 3, 1, 1)

    Combox5 = create_combobox(parameters["part_resistance"][-1], 0, 0)
    widgets["selected_box5"].append(Combox5)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box5"][-1], 2, 4, 1, 1)

    #9th
    label7 = create_label("廠商", 0, 0)
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 3, 3, 1, 1)

    Combox6 = create_combobox(parameters["part_manufacturer"][-1], 0, 0)
    widgets["selected_box6"].append(Combox6)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box6"][-1], 3, 4, 1, 1)

    #10th 11th
    label8 = create_label("供應商", 0, 0)
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 4, 3, 1, 1)

    Combox7 = create_combobox(parameters["part_supplier"][-1], 0, 0)
    widgets["selected_box7"].append(Combox7)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box7"][-1], 4, 4, 1, 1)

    #button widget
    button2 = create_button("生成料號", "#FFFDD4", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_output"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button_output"][-1], 7, 1, 1, 1)

    #LineEdit widget
    lineEdit2 = create_lineedit(0,0)
    lineEdit2.setReadOnly(True)
    widgets["line_bar2"].append(lineEdit2)
    grid.addWidget(widgets["line_bar2"][-1], 7, 2, 1, 3)
    
    label9 = create_label("品項名稱", 0, 0)
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 8, 1, 1, 1)

    #LineEdit widget
    lineEdit3 = create_lineedit(0,0)
    lineEdit3.setReadOnly(True)
    widgets["line_bar3"].append(lineEdit3)
    grid.addWidget(widgets["line_bar3"][-1], 8, 2, 1, 3)
    

    label10 = create_label("已產生編號", 0, 0)
    widgets["label10"].append(label10)
    grid.addWidget(widgets["label10"][-1], 2, 6, 1, 1)

    #QBoxLayout for produced part
    # label11 = create_label("", 0, 0)
    # label11.setStyleSheet("background: #DFDFDF")
    # widgets["label11"].append(label11)
    # grid.addWidget(widgets["label11"][-1], 3, 6, 4, 1)

    grid.addWidget(Pannel, 3, 6, 4, 1)
    # grid.addLayout(vbox_trash_layout, 3, 6, 4, 1)

    button3 = create_button("匯出至Excel", "#1F7145", 0, 0)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 7, 6, 1, 1)

    button2.clicked.connect(lambda: part_number_generate_for_frame4(lineEdit2, lineEdit3, widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText(), widgets["selected_box3"][-1].currentText(), widgets["selected_box4"][-1].currentText(), widgets["selected_box5"][-1].currentText(), widgets["selected_box6"][-1].currentText(), widgets["selected_box7"][-1].currentText()))
    button3.clicked.connect(export_data_to_excel)

#*********************************************
#                  FRAME 5
#*********************************************

def frame5(part_choose, type_choose):
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    #import combo box data
    parameters["part_object"].append(object_selected)

    parameters["part_type"].append([list(d.keys())[0] for d in type_data[object_selected[part_choose]]])

    #info widget
    label1 = create_label("料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 1, 1, 1, 1)

     #LineEdit widget
    lineEdit = create_lineedit(0,0)
    widgets["line_bar1"].append(lineEdit)
    grid.addWidget(widgets["line_bar1"][-1], 1, 2, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#008E8E", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 1, 4, 1, 1)

    #info widget
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 2, 1, 1, 1)

    #button widget
    Combox1 = create_combobox(parameters["part_object"][-1], 0, 0)
    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_object_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 2, 2, 1, 1)

    #info widget
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 3, 1, 1, 1)

    #button widget
    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].setCurrentIndex(type_choose)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 3, 2, 1, 1)

    try:
        parameters["part_kind"].append([list(d.keys())[0] for d in kind_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_kind"].append(["貼片IC"])

    #info widget
    label4 = create_label("種類", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 4, 1, 1, 1)
    #button widget
    Combox3 = create_combobox(parameters["part_kind"][-1], 0, 0)
    widgets["selected_box3"].append(Combox3)
    widgets["selected_box3"][-1].currentIndexChanged.connect(lambda: combo_part_kind_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box3"][-1].currentIndex(), widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText(), widgets["selected_box3"][-1].currentText()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box3"][-1], 4, 2, 1, 1)

    label11 = create_label("已產生編號", 0, 0)
    widgets["label11"].append(label11)
    grid.addWidget(widgets["label11"][-1], 2, 6, 1, 1)

    grid.addWidget(Pannel, 3, 6, 4, 1)

    button3 = create_button("匯出至Excel", "#1F7145", 0, 0)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 7, 6, 1, 1)

#*********************************************
#                  FRAME 6
#*********************************************

def frame6(part_choose, type_choose, kind_choose):
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    #import combo box data
    parameters["part_object"].append(object_selected)

    parameters["part_type"].append([list(d.keys())[0] for d in type_data[object_selected[part_choose]]])

    #info widget
    label1 = create_label("料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 1, 1, 1, 1)

     #LineEdit widget
    lineEdit = create_lineedit(0,0)
    widgets["line_bar1"].append(lineEdit)
    grid.addWidget(widgets["line_bar1"][-1], 1, 2, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#008E8E", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 1, 4, 1, 1)

    #info widget
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 2, 1, 1, 1)

    #button widget
    Combox1 = create_combobox(parameters["part_object"][-1], 0, 0)
    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_object_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 2, 2, 1, 1)

    #info widget
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 3, 1, 1, 1)

    #button widget
    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].setCurrentIndex(type_choose)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 3, 2, 1, 1)

    try:
        parameters["part_kind"].append([list(d.keys())[0] for d in kind_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_kind"].append(["貼片IC"])

    #info widget
    label4 = create_label("種類", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 4, 1, 1, 1)
    #button widget
    Combox3 = create_combobox(parameters["part_kind"][-1], 0, 0)
    widgets["selected_box3"].append(Combox3)
    widgets["selected_box3"][-1].setCurrentIndex(kind_choose)
    widgets["selected_box3"][-1].currentIndexChanged.connect(lambda: combo_part_kind_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box3"][-1].currentIndex(), widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText(), widgets["selected_box3"][-1].currentText()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box3"][-1], 4, 2, 1, 1)

    try:
        parameters["part_name"].append([list(d.keys())[0] for d in name_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_name"].append(["RS621KXF", "LTL431ALT1G(乐山)", "ID5S609SEC-R1", "ULN2003G(UTC)"])
    
    try:
        parameters["part_manufacturer"].append([list(d.keys())[0] for d in manufacturer_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_manufacturer"].append([])

    try:
        parameters["part_supplier"].append([list(d.keys())[0] for d in supplier_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_supplier"].append([])

    #info widget
    label5 = create_label("料件名稱", 0, 0)
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 5, 1, 1, 1)

    Combox4 = create_combobox(parameters["part_name"][-1], 0, 0)
    widgets["selected_box4"].append(Combox4)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box4"][-1], 5, 2, 1, 1)

    #9th
    label6 = create_label("廠商", 0, 0)
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 4, 3, 1, 1)

    Combox5 = create_combobox(parameters["part_manufacturer"][-1], 0, 0)
    widgets["selected_box5"].append(Combox5)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box5"][-1], 4, 4, 1, 1)

    #10th 11th
    label7 = create_label("供應商", 0, 0)
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 5, 3, 1, 1)

    Combox6 = create_combobox(parameters["part_supplier"][-1], 0, 0)
    widgets["selected_box6"].append(Combox6)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box6"][-1], 5, 4, 1, 1)

    #button widget
    button2 = create_button("生成料號", "#FFFDD4", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_output"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button_output"][-1], 7, 1, 1, 1)

    #LineEdit widget
    lineEdit2 = create_lineedit(0,0)
    lineEdit2.setReadOnly(True)
    widgets["line_bar2"].append(lineEdit2)
    grid.addWidget(widgets["line_bar2"][-1], 7, 2, 1, 3)

    label8 = create_label("品項名稱", 0, 0)
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 8, 1, 1, 1)

    #LineEdit widget
    lineEdit3 = create_lineedit(0,0)
    widgets["line_bar3"].append(lineEdit3)
    grid.addWidget(widgets["line_bar3"][-1], 8, 2, 1, 3)

    label9 = create_label("已產生編號", 0, 0)
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 2, 6, 1, 1)

    #QBoxLayout for produced part
    # label10 = create_label("", 0, 0)
    # label10.setStyleSheet("background: #DFDFDF")
    # widgets["label10"].append(label10)
    # grid.addWidget(widgets["label10"][-1], 3, 6, 4, 1)

    grid.addWidget(Pannel, 3, 6, 4, 1)
    # grid.addLayout(vbox_trash_layout, 3, 7, 4, 1)

    button3 = create_button("匯出至Excel", "#1F7145", 0, 0)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 7, 6, 1, 1)

    button2.clicked.connect(lambda: part_number_generate_for_frame6(lineEdit2, widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText(), widgets["selected_box3"][-1].currentText(), widgets["selected_box4"][-1].currentText(), widgets["selected_box5"][-1].currentText(), widgets["selected_box6"][-1].currentText()))
    button3.clicked.connect(export_data_to_excel)

#*********************************************
#                  FRAME 7
#*********************************************

def frame7(part_choose, type_choose, kind_choose):
    global parameters

    clear_widgets()

    clear_parameters(parameters)

    #import combo box data
    parameters["part_object"].append(object_selected)

    parameters["part_type"].append([list(d.keys())[0] for d in type_data[object_selected[part_choose]]])

    #info widget
    label1 = create_label("料件資訊", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 1, 1, 1, 1)

     #LineEdit widget
    lineEdit = create_lineedit(0,0)
    widgets["line_bar1"].append(lineEdit)
    grid.addWidget(widgets["line_bar1"][-1], 1, 2, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#008E8E", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_input"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button_input"][-1], 1, 4, 1, 1)

    #info widget
    label2 = create_label("項目", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 2, 1, 1, 1)

    #button widget
    Combox1 = create_combobox(parameters["part_object"][-1], 0, 0)
    widgets["selected_box1"].append(Combox1)
    widgets["selected_box1"][-1].setCurrentIndex(part_choose)
    widgets["selected_box1"][-1].currentIndexChanged.connect(lambda: combo_part_object_change(widgets["selected_box1"][-1].currentIndex()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box1"][-1], 2, 2, 1, 1)

    #info widget
    label3 = create_label("種類", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 3, 1, 1, 1)

    #button widget
    Combox2 = create_combobox(parameters["part_type"][-1], 0, 0)
    widgets["selected_box2"].append(Combox2)
    widgets["selected_box2"][-1].setCurrentIndex(type_choose)
    widgets["selected_box2"][-1].currentIndexChanged.connect(lambda: combo_part_type_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box2"][-1], 3, 2, 1, 1)

    kind_data = execl_handle.execl_size_read(df, object_selection)
    try:
        parameters["part_kind"].append([list(d.keys())[0] for d in kind_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_kind"].append(["貼片IC"])

    #info widget
    label4 = create_label("種類", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 4, 1, 1, 1)
    #button widget
    Combox3 = create_combobox(parameters["part_kind"][-1], 0, 0)
    widgets["selected_box3"].append(Combox3)
    widgets["selected_box3"][-1].setCurrentIndex(kind_choose)
    widgets["selected_box3"][-1].currentIndexChanged.connect(lambda: combo_part_kind_change(widgets["selected_box1"][-1].currentIndex(), widgets["selected_box2"][-1].currentIndex(), widgets["selected_box3"][-1].currentIndex(), widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText(), widgets["selected_box3"][-1].currentText()))

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box3"][-1], 4, 2, 1, 1)

    try:
        parameters["part_percentage"].append([list(d.keys())[0] for d in percentage_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_percentage"].append([])

    try:
        parameters["part_name"].append([list(d.keys())[0] for d in name_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_name"].append(["RS621KXF", "LTL431ALT1G(乐山)", "ID5S609SEC-R1", "ULN2003G(UTC)"])
    
    try:
        parameters["part_voltage"].append([list(d.keys())[0] for d in voltage_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_voltage"].append(["16V", "25V", "50V", "100V"])
    
    try:
        parameters["part_manufacturer"].append([list(d.keys())[0] for d in manufacturer_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_manufacturer"].append([])

    try:
        parameters["part_supplier"].append([list(d.keys())[0] for d in supplier_data[widgets["selected_box2"][-1].currentText()]])
    except:
        parameters["part_supplier"].append([])

    #info widget
    label5 = create_label("%數", 0, 0)
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 5, 1, 1, 1)

    Combox4 = create_combobox(parameters["part_percentage"][-1], 0, 0)
    widgets["selected_box4"].append(Combox4)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box4"][-1], 5, 2, 1, 1)

    #info widget
    label6 = create_label("料件名稱", 0, 0)
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 2, 3, 1, 1)

    Combox5 = create_combobox(parameters["part_name"][-1], 0, 0)
    widgets["selected_box5"].append(Combox5)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box5"][-1], 2, 4, 1, 1)

    #info widget
    label7 = create_label("電壓", 0, 0)
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 3, 3, 1, 1)

    Combox6 = create_combobox(parameters["part_voltage"][-1], 0, 0)
    widgets["selected_box6"].append(Combox6)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box6"][-1], 3, 4, 1, 1)

    #9th
    label8 = create_label("廠商", 0, 0)
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 4, 3, 1, 1)

    Combox7 = create_combobox(parameters["part_manufacturer"][-1], 0, 0)
    widgets["selected_box7"].append(Combox7)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box7"][-1], 4, 4, 1, 1)

    #10th 11th
    label9 = create_label("供應商", 0, 0)
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 5, 3, 1, 1)

    Combox8 = create_combobox(parameters["part_supplier"][-1], 0, 0)
    widgets["selected_box8"].append(Combox8)

    #place global widgets on the grid
    grid.addWidget(widgets["selected_box8"][-1], 5, 4, 1, 1)

    #button widget
    button2 = create_button("生成料號", "#FFFDD4", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button_output"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button_output"][-1], 7, 1, 1, 1)

    #LineEdit widget
    lineEdit2 = create_lineedit(0,0)
    lineEdit2.setReadOnly(True)
    widgets["line_bar2"].append(lineEdit2)
    grid.addWidget(widgets["line_bar2"][-1], 7, 2, 1, 3)

    label10 = create_label("品項名稱", 0, 0)
    widgets["label10"].append(label10)
    grid.addWidget(widgets["label10"][-1], 8, 1, 1, 1)

    #LineEdit widget
    lineEdit3 = create_lineedit(0,0)
    widgets["line_bar3"].append(lineEdit3)
    grid.addWidget(widgets["line_bar3"][-1], 8, 2, 1, 3)

    label11 = create_label("已產生編號", 0, 0)
    widgets["label11"].append(label11)
    grid.addWidget(widgets["label11"][-1], 2, 6, 1, 1)

    #QBoxLayout for produced part
    # label12 = create_label("", 0, 0)
    # label12.setStyleSheet("background: #DFDFDF")
    # widgets["label12"].append(label12)
    # grid.addWidget(widgets["label12"][-1], 3, 6, 4, 1)

    grid.addWidget(Pannel, 3, 6, 4, 1)
    # grid.addLayout(vbox_trash_layout, 3, 6, 4, 1)

    button3 = create_button("匯出至Excel", "#1F7145", 0, 0)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 7, 6, 1, 1)

    button2.clicked.connect(lambda: part_number_generate_for_frame7(lineEdit2, widgets["selected_box1"][-1].currentText(), widgets["selected_box2"][-1].currentText(), widgets["selected_box3"][-1].currentText(), widgets["selected_box4"][-1].currentText(), widgets["selected_box5"][-1].currentText(), widgets["selected_box6"][-1].currentText(), widgets["selected_box7"][-1].currentText(),  widgets["selected_box8"][-1].currentText()))
    button3.clicked.connect(export_data_to_excel)
    
    