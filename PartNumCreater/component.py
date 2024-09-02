from PyQt5.QtWidgets import QHBoxLayout, QTreeWidget, QTreeWidgetItem, QTableWidgetItem, QWidget, QGridLayout, QScrollArea, QLabel, QPushButton, QLineEdit, QComboBox, QDateEdit, QDialog, QVBoxLayout, QGraphicsDropShadowEffect, QTableWidget, QInputDialog, QTabWidget, QTextEdit, QCheckBox
from PyQt5.QtGui import QCursor, QColor, QPixmap
from PyQt5 import QtCore
from PyQt5.QtCore import Qt, QDate
from printer import tag_printer
from tag_pic_creater import draw_tag_sticker
import execl_handle, SQL_handler
import os
import pandas as pd
from datetime import datetime
import Name_Rule_SQL_handler
import pyperclip



from_print_page = 0
from_generate_page = 1

out_stock_count = 0

grid = QGridLayout()
main_grid = QGridLayout()

search_page_widgets = {
    "button_back":[],
    "label0":[], #input
    "label1":[], #part number
    "label2":[], #part code
    "label3":[], #part name
    "label4":[], #part object
    "label5":[], #part type
    "label6":[], #part size/kind
    "label7":[], #part percentage
    "label8":[], #part capacity/resistance/name
    "label9":[], #part voltage
    "label10":[], #part manufacturer
    "label11":[], #part supplier
    "label12":[],  #part generate date
    "label13":[],  #part generate date
    "label14":[],  #part generate date
    "label15":[],  #part generate date
    "label16":[],  #part generate date
    "label17":[],  #part generate date
    "label_pn":[],
    "line_bar0":[], #input number
    "line_bar1":[], #part number
    "line_bar2":[], #part code
    "line_bar3":[], #part name
    "line_bar4":[], #part object
    "line_bar5":[], #part type
    "line_bar6":[], #part size/kind
    "line_bar7":[], #part percentage
    "line_bar8":[], #part capacity/resistance/name
    "line_bar9":[], #part voltage
    "line_bar10":[], #part manufacturer
    "line_bar11":[], #part supplier
    "line_bar12":[],  #part generate date
    "line_bar13":[],  #part generate date
    "line_bar14":[],  #part generate date
    "line_bar15":[],  #part generate date
    "line_bar16":[],  #part generate date
    "line_bar17":[],  #part generate date
    "line_bar_pn":[], 
    "button":[], #part code input
    "button_search":[], #part code input
}

#global dictionary of dynamically changing widgets
generate_page_widgets = {
    "button_back":[],
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
    "label_w":[],
    "label_pn":[],
    "line_bar1": [],
    "line_bar2": [],
    "line_bar3": [],
    "line_bar_pn": [],
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
    "selected_w":[],
}

barcode_page_widgets={
    "button_back":[],
    "button1":[],
    "button2":[],
    "button3":[],
    "button4":[],
    "button_printer":[],
    "button_sure":[],
    "button_input":[],
    "button_print":[],
    "button_print_export":[],
    "button_operation":[],
    "button_delete":[],
    "button_ecount":[],
    "button_single_out":[],
    "button_batch_out":[],
    "label1":[],
    "label2":[],
    "label3":[],
    "label4":[],
    "label5":[],
    "label6":[],
    "label7":[],
    "label8":[],
    "label9":[],
    "label_remand":[],
    "label_box_num":[],
    "label_in_date":[],
    "label_rollover_area":[],
    "label_rollover_place":[],
    "label_box":[],
    "line_bar1":[],
    "line_bar2":[],
    "line_bar3":[],
    "line_bar_box_num":[],
    "line_bar_box":[],
    "line_bar_hide":[],
    "line_bar_batch":[],
    "date_choose1":[],
    "date_choose2":[],
    "date_in":[],
    "combo1":[],
    "combo2":[],
    "combo3":[],
    "combo_rollover_area":[],
    "combo_rollover_place":[],
    "Tab": [],
    "text_edit": [],
    "table_print":[],
    "table_show":[],
    "chk_box": [],
    "lineedit_cover":[],
    "lineedit_cover2":[],
}

main_page_widgets = {
    "title":[], 
    "button1":[],
    "button2":[], 
    "button3":[],
    "button4":[],
}

login_page_widgets = {
    "label_account":[],
    "label_password":[],
    "label_password2":[],
    "label_message":[],
    "label_message2":[],
    "label_mail":[],
    "line_account":[],
    "line_password":[],
    "line_password2":[],
    "line_mail":[],
    "button_login":[],
    "button_signin":[],
    "button_forget_password":[],
    "button_eye":[],
    "button_send_mail":[],
    "line_certify":[],
    "label_certify":[],
}

material_shortage_widgets = {
    "button_save_forecast":[],
    "button_delete_forecast":[],
    "button_back":[],
    "button_import":[],
    "button_load_excel":[],
    "button_calculate":[],
    "button_export":[],
    "button_eye":[],
    "linedit_filename":[],
    "table":[],
    "table2":[],
    "navy_bar":[],
    "button_hide_and_show":[],
    "label_or":[],
    "combo_forecast":[],
}

part_code_generate_page_widget = {
    "label_item": [],
    "label_category": [],
    "label_type": [],
    "label11": [],
    "label_pn": [],
    "combo_item": [],
    "combo_category": [],
    "combo_type": [],
    "button1": [],
    "button2": [],
    "button3": [],
    "button_back": [],
    "button_output": [],
    "button_export": [],
    "button_insert": [],
    "button_delete": [],
    "button_show":[],
    "line_bar2": [],
    "line_bar3": [],
    "line_bar_pn":[],
    "label12": [],
    "label_choose": [],
    "combo_choose": [],
    "table": [],
}

part_code_generate_page_hide_widget = {
    "label_NUM1":[],
    "label_NUM2":[],
    "label_NUM3":[],
    "label_NUM4":[],
    "label_NUM5":[],
    "label_NUM6":[],
    "label_NUM7":[],
    "label_NUM8":[],
    "label_NUM9":[],
    "label_NUM10":[],
    "label_NUM11":[],
    "label_NUM12":[],
    "combo_NUM1":[],
    "combo_NUM2":[],
    "combo_NUM3":[],
    "combo_NUM4":[],
    "combo_NUM5":[],
    "combo_NUM6":[],
    "combo_NUM7":[],
    "combo_NUM8":[],
    "combo_NUM9":[],
    "combo_NUM10":[],
    "combo_NUM11":[],
    "combo_NUM12":[],
}

def clear_widgets(widgets):
    ''' hide all existing widgets and erase
        them from the global dictionary'''
    for widget in widgets:
        if widgets[widget] != []:
            widgets[widget][-1].hide()
        for i in range(0, len(widgets[widget])):
            widgets[widget].pop()

def create_label(name, l_margin, r_margin, align = 'left', width = 150, height = 44, font_size = 20):
    #create identical buttons with custom left & right margins
    label = QLabel(name)
    label.setFixedWidth(width)
    label.setFixedHeight(height)
    label.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        "font-size: " + str(font_size) +"px;"
        '''
            color: 'black';
            padding: 0px 0px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
            background: rgba(0,0,0,0);
        }
        '''
    )
    if(align == 'center'):
        label.setAlignment(Qt.AlignVCenter | Qt.AlignHCenter)
    elif(align == 'right'):
        label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
    return label

def create_lineedit(l_margin, r_margin, width = 300, font_size = 20):
    lineedit = QLineEdit()
    lineedit.setFixedWidth(width)
    lineedit.setFixedHeight(44)
    lineedit.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        "font-size: " + str(font_size) +"px;"+
        '''
            background: '#D9D9D9';
            color: 'black';
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        *:hover{
            background: '#ECECEC';
            font-size: 20px;
            color: 'black';
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        '''
    )
    return lineedit

def create_dateedit(l_margin, r_margin, width = 300, font_size = 20):
    dateedit = QDateEdit()
    dateedit.setFixedWidth(width)
    dateedit.setFixedHeight(44)
    dateedit.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        "font-size: " + str(font_size) +"px;"
        '''
            background: '#D9D9D9';
            color: 'black';
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        *:hover{
            background: '#ECECEC';
            font-size: 20px;
            color: 'black';
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        '''
    )
    dateedit.setCalendarPopup(True)
    dateedit.setDate(QDate.currentDate())

    return dateedit

def create_button(name, back_color ,l_margin, r_margin, height = 44, width = 110, qcolor = [150,150,150, 160],bot_margin=0, font=20, font_color= 'black'):
    button = QPushButton(name)
    button.setFixedWidth(width)
    button.setFixedHeight(height)
    button.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
    button.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        "margin-bottom: " + str(bot_margin) +"px;"+
        "background: " + back_color +";"+
        "font-size: " + str(font) +"px;"+
        "color: " + font_color +";"+
        '''
            padding: 0px 0px;
            font-family: 'Microsoft JhengHei';
            font-weight: bold;
        }
        *:hover{
        '''
            "background: rgba(" + str(qcolor[0]) + "," + str(qcolor[1]) + "," + str(qcolor[2]) + "," + str(qcolor[3]) + ");"
        '''
        }
        '''
    )
    # border-radius: 10px;
    # button.clicked.connect(start_game)
    # shadow_effect = QGraphicsDropShadowEffect()
    # shadow_effect.setBlurRadius(10)
    # shadow_effect.setXOffset(5)
    # shadow_effect.setYOffset(5)
    # shadow_effect.setColor(QColor(qcolor[0], qcolor[1], qcolor[2], qcolor[3]))
    # button.setGraphicsEffect(shadow_effect)
    return button

def create_table():
    table = QTableWidget()
    # table.setFixedWidth(width)
    table.setStyleSheet(
        """
        QTableWidget {
            background-color: #F0F0F0;
            alternate-background-color: #D5D5D5;
            gridline-color: #000000;
            font-size: 14px;
            font-family: Microsoft JhengHei;
        }
        QHeaderView::section {
            background-color: #404040;
            color: #FFFFFF;
            padding: 4px;
            border: 1px solid #6c6c6c;
        }
        QTableWidget::item {
            padding: 5px;
            border: 1px solid #6c6c6c;
        }
        QTableWidget::item:selected {
            background-color: #3399FF;
            color: #FFFFFF;
        }
    """
    )
    table.setAlternatingRowColors(True)

    return table

# def button_clicked(button, color, l_margin, r_margin, qcolor):
#     button.setStyleSheet(
#         #setting variable margins
#         "*{margin-left: " + str(l_margin) +"px;"+
#         "margin-right: " + str(r_margin) +"px;"+
#         "background: " + color +";"+
#         '''
#             font-size: 20px;
#             color: 'black';
#             padding: 0px 0px;
#             border-radius: 10px;
#             font-family: 'Microsoft JhengHei';
#             font-weight: bold;
#         }
#         '''
#     )

#     # button.clicked.connect(start_game)
#     shadow_effect = QGraphicsDropShadowEffect()
#     shadow_effect.setBlurRadius(10)
#     shadow_effect.setXOffset(5)
#     shadow_effect.setYOffset(5)
#     shadow_effect.setColor(QColor(qcolor[0], qcolor[1], qcolor[2], qcolor[3]))
#     button.setGraphicsEffect(shadow_effect)

def create_combobox(item_list, l_margin, r_margin, width= 150, font_size = 20):
    combo = QComboBox()
    combo.setFixedWidth(width)
    combo.setFixedHeight(44)
    combo.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
    combo.addItems(item_list)
    combo.setCurrentIndex(-1)
    combo.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        "font-size: " + str(font_size) +"px;"+
        '''
            background: '#D9D9D9';
            color: 'black';
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        *:hover{
            background: '#ECECEC';
        }
        '''
    )

    return combo

def create_tab(l_margin, r_margin, width= 20, font_size = 15):
    tab = QTabWidget()
    tab.setStyleSheet (
        "QTabBar::tab{padding-left: " + str(l_margin) +"px;"+
        "padding-right: " + str(r_margin) +"px;"+
        "font-size: " + str(font_size) +"px;"+
        '''
            height: 30px;  /* 设置标签的高度 */
            width: 40px;  /* 设置标签的宽度 */
            color: 'black';
            padding: 0px 0px;
            border-radius: 10px;
            font-family: 'Microsoft JhengHei';
            font-weight: bold;
        }
        QTabBar::tab:selected {
            background: #F7DED0;  /* 设置选中标签的背景颜色 */
            color: #FFFFFF;  /* 设置选中标签的文本颜色 */
        }
        '''
    )
    
    return tab

def create_textEdit(l_margin, r_margin,  font_size = 15):
    textedit = QTextEdit()
    # textedit.setFixedWidth(200)
    textedit.setStyleSheet(
        "*{padding-left: " + str(l_margin) +"px;"+
        "padding-right: " + str(r_margin) +"px;"+
        "font-size: " + str(font_size) +"px;"+
        '''
            background: '#D9D9D9';
            color: 'black';
            padding: 0px 0px;
            border-radius: 10px;
            font-family: 'Microsoft JhengHei';
            font-weight: bold;
        }
        '''
    )
    textedit.setEnabled(False)

    return textedit

def create_checkBox(name=""):
    checkbox = QCheckBox(name)
    checkbox.setStyleSheet(
        """
        QCheckBox {
            font-size: 20px;
            font-weight: bold;
            font-family: 'Microsoft JhengHei';
            margin-left:45%; 
            margin-right:45%;
        }
        QCheckBox::indicator { 
            width: 20px; 
            height: 20px;
        }
        """
    )

    return checkbox

def create_checkBox_header():
    checkbox = QCheckBox()
    checkbox.setFixedHeight(80)
    checkbox.setFixedWidth(120)
    checkbox.setStyleSheet(
        """
        QCheckBox {
            margin-left:90%; 
            margin-right:0%;
            margin-bottom:20%;
            margin-top:70%;
        }
        QCheckBox::indicator { 
            width: 20px; 
            height: 20px;
        }
        """
    )
    return checkbox

def show_alert(message, type="警告"):
    dialog = QDialog()
    
    dialog.setWindowTitle(type)
    dialog.resize(200,100)
    layout = QVBoxLayout()
    label = QLabel(message)
    layout.addWidget(label)
    if (type == "警告"): 
        label.setStyleSheet("color: red;")
    elif (type == "通知"):
        label.setStyleSheet("color: green;")

    close_button = QPushButton("Close")
    close_button.clicked.connect(dialog.close)
    layout.addWidget(close_button)
    
    dialog.setLayout(layout)
    
    dialog.exec_()

# 定義函數根據part_number返回整列資訊
def get_part_info(df ,part_number):
    part_info = df[df['ERP Code'] == part_number]
    if not part_info.empty:
        return part_info.values.tolist()
    else:
        return None

def close(dialog, num):
    for i in range(num):
        os.remove("sticker/pic"+str(i)+".png")
    dialog.close()

def save_and_print(inventory_list, boxes_list, line_edit1, dialog, num, code_list, count_list):
    in_update(inventory_list, boxes_list, line_edit1)
    # print(f"inventory_list:{inventory_list}")
    # print(f"boxes_list: {boxes_list}")
    tprint(dialog, num, code_list, boxes_list)

def tprint(dialog, num, code_list, count_list, page_from = 1):
    
    for idx,val in enumerate(count_list):
        try:
            tag_printer("sticker/pic"+str(idx)+".png")
            os.remove("sticker/pic"+str(idx)+".png")
            
        except Exception as e:
            show_alert("列印失敗! 請檢查Brother標籤機")
            return
        
        else:
            # print(f"入庫號碼:{val[0]}, {type(val[0])}\n入庫箱號: {val[1]}, {type(val[1])}")
            box_id = val[0] + execl_handle.number_to_letter(int(val[1]))
            SQL_handler.increment_print_count_by_box_id(box_id)

       
    dialog.close()
    if(page_from == from_print_page):
        from barcode_generator import print_page
        print_page()

def preview2(num, code_list, count_list):
    dialog = QDialog()
    dialog.setWindowTitle("預覽列印")
    dialog.resize(550,550)
    layout = QVBoxLayout()

    scroll_area = QScrollArea()
    scroll_area_widget_contents = QWidget()
    scroll_layout = QVBoxLayout(scroll_area_widget_contents)

    for i in range(num):
        pixmap = QPixmap("sticker/pic"+str(i)+".png")
        label = QLabel()
        label.setPixmap(pixmap)
        scroll_layout.addWidget(label)

    scroll_area.setWidget(scroll_area_widget_contents)
    layout.addWidget(scroll_area)

    close_button = QPushButton("Close")
    close_button.clicked.connect(lambda: close(dialog, num))
    layout.addWidget(close_button)

    print_button = QPushButton("Print")
    print_button.clicked.connect(lambda: tprint(dialog, num, code_list, count_list, from_print_page))
    layout.addWidget(print_button)
    
    dialog.setLayout(layout)
    
    dialog.exec_()

def preview3(num, part_code, count, in_stock_time, stock_place , line_bar2, check_code: int , line_edit1,table_show,date_in):

    if(num == ''): 
        show_alert("請選擇箱數")
        return
    else: num = int(num)

    total_num = int(line_bar2.text())

    global out_stock_count
    
    if(check_code == 0):
        check = 0
        for i in range(num):
            try:
                check = check + int(count['box'+str(i)][-1].text())
            except:
                show_alert("請確實輸入每筆數量")
                return
        if(total_num != check):
            show_alert("數量加總與總數量不同!")
            return
        # count = total_num
    elif(check_code == 1):
        if(count > total_num):
            show_alert("出庫數量不可大於庫存數量!")
            return
        
        count = total_num - count
    elif(check_code == 2):
        if(out_stock_count == 0):
            
            out_stock_count = count

        if(count > out_stock_count):
            show_alert("退庫數量不可大於出庫數量!")
            return
        
        count = total_num + count 

    output_list = SQL_handler.fetch_data_from_Material_table({"ERP_Code":part_code})

    if (output_list == []):
        show_alert("零件編碼未建立!")
    else:
        _list = output_list[0]
        if (str(_list[4]) == "電容"):
            part_spec = str(_list[3])+", "+str(_list[5]).zfill(4)+", "+ str(_list[2]).split(',')[0] +", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
            part_name = "電容/電容"
        elif (str(_list[4]) == "電阻"):
            part_spec = str(_list[3])+", "+str(_list[5]).zfill(4)+", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
            part_name = "電阻/電阻"
        elif ("電容" in str(_list[4])):
            part_spec = str(_list[3])+", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
            part_name = str(_list[4])+"/"+str(_list[5])
        elif ("機構元件" in str(_list[4]) or "晶振" in str(_list[4])):
            part_spec = str(_list[3])+", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
            part_name = str(_list[4])+"/"+str(_list[5])
        elif("PCB" in str(_list[4])):
            part_spec = str(_list[3])+"."+str(_list[8])+", "+str(_list[7])
            part_name = str(_list[4])+"/"+str(_list[5])
        else:
            part_spec = str(_list[3])+", "+str(_list[7]).split('(')[0].split('/')[0]
            part_name = str(_list[4])+"/"+str(_list[5])

        manufacter = _list[9]
        supplier = _list[10]
        pn = str(_list[12])

        dialog = QDialog()
        dialog.setWindowTitle("預覽列印")
        dialog.resize(550,550)
        layout = QVBoxLayout()

        scroll_area = QScrollArea()
        scroll_area_widget_contents = QWidget()
        scroll_layout = QVBoxLayout(scroll_area_widget_contents)

        
        inventory_list= []
        boxes_list= []
        from barcode_generator import placecode, stockroom
        # 取得倉庫編號以及對應箱號
        inventory_id = part_code+in_stock_time.replace('-','')[2:]+str(placecode[stockroom[stock_place][0]])+str(stockroom[stock_place][1])
        # print("產生的入庫資訊 ",inventory_id)
        box_num = SQL_handler.get_max_serial_number(inventory_id)
        from longin_page import user
        
        inventory_list = [inventory_id, part_code, total_num, stock_place, in_stock_time, user]

        code_list = []
        count_list = []

        create_time = datetime.now().strftime('%Y-%m-%d')
    
        for i in range(num):
            # boxes_list.append(draw_tag_sticker("pic"+str(i), part_code, part_spec, manufacter, supplier, int(count['box'+str(i)][-1].text()), part_name, in_stock_time, create_time, stock_place, pn, box_num+i))
            code_list.append(part_code)
            count_list.append(box_num+i)
            if(check_code == 0):
                label = QLabel()
                # print("create_PIC")
                boxes_list.append(draw_tag_sticker("sticker/pic"+str(i), part_code, part_spec, manufacter, supplier, int(count['box'+str(i)][-1].text()), part_name, in_stock_time, create_time, stock_place, pn, box_num+i, count['batch_box'+str(i)][-1].text(), stock_place))
                pixmap = QPixmap("sticker/pic"+str(i)+".png")
                label.setPixmap(pixmap)
                scroll_layout.addWidget(label)
            else:
                label = QLabel()
                draw_tag_sticker("sticker/pic"+str(i), part_code, part_spec, manufacter, supplier, count, part_name, in_stock_time, create_time, stock_place, pn)
                pixmap = QPixmap("sticker/pic"+str(i)+".png")
                label.setPixmap(pixmap)
                scroll_layout.addWidget(label)

                if (check_code == 1):
                    label = QLabel()
                    draw_tag_sticker("sticker/pic"+str(i)+"-o", part_code, part_spec, manufacter, supplier, total_num - count, part_name, in_stock_time, create_time, stock_place, pn)
                    pixmap = QPixmap("sticker/pic"+str(i)+"-o.png")
                    label.setPixmap(pixmap)
                    scroll_layout.addWidget(label)
            
            
        scroll_area.setWidget(scroll_area_widget_contents)
        layout.addWidget(scroll_area)


        close_button = QPushButton("Close")
        close_button.clicked.connect(lambda: close(dialog, num))
        layout.addWidget(close_button)

        print_button = QPushButton("Print")
        print_button.clicked.connect(lambda: save_and_print(inventory_list, boxes_list, line_edit1, dialog, num, code_list, count_list))
        layout.addWidget(print_button)
        
        dialog.setLayout(layout)
        
        dialog.exec_()

        

        from barcode_generator import in_stock_inform
        in_stock_inform(table_show, date_in.text())

def preview(num, part_code, count, in_stock_time, stock_place , line_bar2, check_code: int , line_edit1,table_show,date_in):
    if(num == ''): 
        show_alert("請選擇箱數")
        return
    else: num = int(num)

    total_num = int(line_bar2.text())

    global out_stock_count
    
    if(check_code == 0):
        check = 0
        for i in range(num):
            try:
                check = check + int(count['box'+str(i)][-1].text())
            except:
                show_alert("請確實輸入每筆數量")
                return
        if(total_num != check):
            show_alert("數量加總與總數量不同!")
            return
        # count = total_num
    elif(check_code == 1):
        if(count > total_num):
            show_alert("出庫數量不可大於庫存數量!")
            return
        
        count = total_num - count
    elif(check_code == 2):
        if(out_stock_count == 0):
            
            out_stock_count = count

        if(count > out_stock_count):
            show_alert("退庫數量不可大於出庫數量!")
            return
        
        count = total_num + count 
    
    if(len(part_code) != 11):
        show_alert("輸入編碼長度錯誤!")
        return

    output_list = SQL_handler.fetch_data_from_Material_table({"ERP_Code":part_code})
    # print(f"output_list: {output_list}")
    if (output_list == []):
        show_alert("零件編碼未建立!")
    else:
        _list = output_list[0]
        if (str(_list[4]) == "電容"):
            part_spec = str(_list[3])+", "+str(_list[5]).zfill(4)+", "+ str(_list[2]).split(',')[0] +", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
            part_name = "電容/電容"
        elif (str(_list[4]) == "電阻"):
            part_spec = str(_list[3])+", "+str(_list[5]).zfill(4)+", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
            part_name = "電阻/電阻"
        elif ("電容" in str(_list[4])):
            part_spec = str(_list[3])+", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
            part_name = str(_list[4])+"/"+str(_list[5])
        elif ("機構元件" in str(_list[4]) or "晶振" in str(_list[4])):
            part_spec = str(_list[3])+", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
            part_name = str(_list[4])+"/"+str(_list[5])
        elif("PCB" in str(_list[4])):
            part_spec = str(_list[3])+"."+str(_list[8])+", "+str(_list[7])
            part_name = str(_list[4])+"/"+str(_list[5])
        else:
            part_spec = str(_list[3])+", "+str(_list[7]).split('(')[0].split('/')[0]
            part_name = str(_list[4])+"/"+str(_list[5])

        manufacter = _list[9]
        supplier = _list[10]
        pn = str(_list[12])

        # dialog = QDialog()
        # dialog.setWindowTitle("預覽列印")
        # dialog.resize(550,550)
        # layout = QVBoxLayout()

        # scroll_area = QScrollArea()
        # scroll_area_widget_contents = QWidget()
        # scroll_layout = QVBoxLayout(scroll_area_widget_contents)

        
        inventory_list= []
        boxes_list= []
        from barcode_generator import placecode, stockroom
        # 取得倉庫編號以及對應箱號
        inventory_id = part_code+in_stock_time.replace('-','')[2:]+str(placecode[stockroom[stock_place][0]])+str(stockroom[stock_place][1])
        # print("產生的入庫資訊 ",inventory_id)
        box_num = SQL_handler.get_max_serial_number(inventory_id)
        # print(f"box_num={box_num}")
        from longin_page import user
        # print(f"user id = {user}")
        inventory_list = [inventory_id, part_code, total_num, stock_place, in_stock_time, user]

        create_time = datetime.now().strftime('%Y-%m-%d')
        # print(create_time)
        for i in range(num):
            boxes_list.append(draw_tag_sticker("sticker/pic"+str(i), part_code, part_spec, manufacter, supplier, int(count['box'+str(i)][-1].text()), part_name, in_stock_time, create_time, stock_place, pn, box_num+i, count['batch_box'+str(i)][-1].text(), stock_place))
        #     if(check_code == 0):
        #         label = QLabel()
        #         boxes_list.append(draw_tag_sticker("pic"+str(i), part_code, part_spec, manufacter, supplier, int(count['box'+str(i)][-1].text()), part_name, in_stock_time, create_time, stock_place, pn, box_num+i))
        #         pixmap = QPixmap("pic"+str(i)+".png")
        #         label.setPixmap(pixmap)
        #         scroll_layout.addWidget(label)
        #     else:
        #         label = QLabel()
        #         draw_tag_sticker("pic"+str(i), part_code, part_spec, manufacter, supplier, count, part_name, in_stock_time, create_time, stock_place, pn)
        #         pixmap = QPixmap("pic"+str(i)+".png")
        #         label.setPixmap(pixmap)
        #         scroll_layout.addWidget(label)

        #         if (check_code == 1):
        #             label = QLabel()
        #             draw_tag_sticker("pic"+str(i)+"-o", part_code, part_spec, manufacter, supplier, total_num - count, part_name, in_stock_time, create_time, stock_place, pn)
        #             pixmap = QPixmap("pic"+str(i)+"-o.png")
        #             label.setPixmap(pixmap)
        #             scroll_layout.addWidget(label)
            
            
        # scroll_area.setWidget(scroll_area_widget_contents)
        # layout.addWidget(scroll_area)


        # close_button = QPushButton("Close")
        # close_button.clicked.connect(lambda: close(dialog, num))
        # layout.addWidget(close_button)

        # print_button = QPushButton("Print")
        # print_button.clicked.connect(lambda: tprint(dialog, inventory_list, boxes_list, check_code))
        # layout.addWidget(print_button)
        
        # dialog.setLayout(layout)
        
        # dialog.exec_()
        in_update(inventory_list, boxes_list, line_edit1)

        from barcode_generator import in_stock_inform
        in_stock_inform(table_show, date_in.text())

def find_stockroom_name(location, index, stockroom):
    for key, value in stockroom.items():
        if value[0] == location and value[1] == index:
            return key
    return None

def get_place_from_code(code, placecode):
    reverse_placecode = {v: k for k, v in placecode.items()}
    return reverse_placecode.get(code, None)

def in_update(inventory_list: list, boxes_list: list, line_edit1, motivation="IN", old_box_id = ""):
    inventory_target = SQL_handler.fetch_data_from_Inventory_table({"InventoryID":inventory_list[0]})
    # print(inventory_target)
    # Check inventory ID was created when SWITCH, if not, create.
    if(isinstance(inventory_target, str)):
        status = SQL_handler.add_data_to_Inventory_table(inventory_list)
        if status != True:
            show_alert(f"建立Inventory ID時出現錯誤: {status}")
            return
    elif(len(inventory_target) == 0):
        status = SQL_handler.add_data_to_Inventory_table(inventory_list)
        if status != True:
            show_alert(f"建立Inventory ID時出現錯誤: {status}")
            return
    else:
        status = SQL_handler.update_total_quantity_if_duplicate(inventory_list[0], inventory_list[2])
        if status != True:
            show_alert(f"更新Inventroy ID {inventory_list[0]} 數量時出現錯誤: {status}")
            return

    for i in range(len(boxes_list)):
        box_id = boxes_list[i][0]+execl_handle.number_to_letter(int(boxes_list[i][1]))
        if (motivation == "IN"):
            status = SQL_handler.add_data_to_Boxes_table([box_id, boxes_list[i][0], boxes_list[i][1], boxes_list[i][2], box_id, boxes_list[i][3], datetime.now(), 0]) # add boxid as unicode while in stock
            # print(f"Box: {SQL_handler.get_table_data('Boxes')}")
        else:
            import time
            time.sleep(1)
            status = SQL_handler.add_data_to_Boxes_table([box_id, boxes_list[i][0], boxes_list[i][1], boxes_list[i][2], old_box_id, boxes_list[i][3], datetime.now(), 0]) # add boxid as unicode while in stock
            # print(f"Box: {SQL_handler.get_table_data('Boxes')}")
        
        if status != True:
            show_alert(f"物料入庫過程異常: {status}")
            #Check total quantity of inventory id in Boxes table. If not equal to Inventory table, adjust total quantity in Inventory table.
            total_quantity_from_Boxes = SQL_handler.get_total_quantity_by_inventory_id(inventory_list[0])
            SQL_handler.update_total_quantity_if_duplicate(inventory_list[0], total_quantity_from_Boxes)
            return
        
        if (motivation == "IN"):
            unicode = SQL_handler.fetch_data_from_Boxes_table({"BoxID":box_id})[0][4]
            from barcode_generator import placecode, stockroom
            if(get_place_from_code(int(box_id[17]),placecode) != None):
                place = get_place_from_code(int(box_id[17]),placecode)
                if(find_stockroom_name(place ,int(box_id[18]), stockroom) != None):
                    destination_stcok_room = find_stockroom_name(place ,int(box_id[18]), stockroom)
            status = SQL_handler.add_data_to_manage_table([unicode, motivation, boxes_list[i][2], None, destination_stcok_room, None, None,datetime.now(), inventory_list[-1]])
        elif (motivation == "SWITCH"):
            unicode = SQL_handler.fetch_data_from_Boxes_table({"BoxID":box_id})[0][4]
            from barcode_generator import placecode, stockroom
            if(get_place_from_code(int(box_id[17]),placecode) != None):
                place = get_place_from_code(int(box_id[17]),placecode)
                if(find_stockroom_name(place ,int(box_id[18]), stockroom) != None):
                    destination_stcok_room = find_stockroom_name(place ,int(box_id[18]), stockroom)
            if(get_place_from_code(int(old_box_id[17]),placecode) != None):
                place = get_place_from_code(int(old_box_id[17]),placecode)
                if(find_stockroom_name(place ,int(old_box_id[18]), stockroom) != None):
                    source_stcok_room = find_stockroom_name(place ,int(old_box_id[18]), stockroom)
            status = SQL_handler.add_data_to_manage_table([unicode, motivation, boxes_list[i][2], source_stcok_room, destination_stcok_room, None, None,datetime.now(), inventory_list[-1]])


        status = SQL_handler.add_data_to_Operation_table([box_id, motivation, 0, boxes_list[i][2], old_box_id,datetime.now(), inventory_list[-1]])
        # print(f"Operation: {SQL_handler.get_table_data('Operation')}")

        if status != True:
            show_alert(f"操作紀錄存入資料庫時發生異常: {status}")
            return
    
    

    show_alert(inventory_list[1]+"料件已更新至資料庫!", "通知")
    line_edit1.clear()
        
def out_update(line_bar2, line_bar_box_num, data1, data2, stock_palce_combo, data4, line_bar3,chk_box_state, combo1, combo2, date_choose2):
    from longin_page import user
    
    if (line_bar3.text() == ''):
        show_alert("請填入出庫數量")
        return
    from barcode_generator import placecode, stockroom

    if (data1.text() == '' or data2.text() == '' or stock_palce_combo.currentText() == '' or data4.text() == ''):
        show_alert("請輸入或掃barcode")
        line_bar3.setText('')
        return
    # BoxID = data1.text()+ data2.text().split('/')[0][2:] + data2.text().split('/')[1].zfill(2)+ data2.text().split('/')[2].zfill(2) +str(placecode[stockroom[stock_palce_combo.currentText()][0]])+str(stockroom[stock_palce_combo.currentText()][1])+ execl_handle.number_to_letter(int(data4.text()))
    BoxID = date_choose2.text().split(',')[1]
    # back_list = execl_handle.out_update_quantity(BoxID, int(line_bar3.text()))
    # print(f"back_list: {back_list}") [50,50]

    if (combo2.currentText() == ""):
            show_alert("請選擇轉倉倉庫")
            return

    Box_info = SQL_handler.fetch_data_from_Boxes_table({"BoxID":BoxID})
    
    if isinstance(Box_info, str):
        show_alert(f"讀取庫存資料異常: {Box_info}")
        return 
    elif Box_info == []:
        show_alert("資料庫Boxes表內無此資料")
        return

    current_quantity = Box_info[0][3]
    # input quantity is larger than quantity in stock
    if int(line_bar3.text()) > current_quantity:
        show_alert(f"Quantity to take ({int(line_bar3.text())}) is greater than current quantity ({current_quantity}).")
        return

    new_quantity = current_quantity - int(line_bar3.text())
    
    status = SQL_handler.update_quantity_by_box_id(BoxID, new_quantity, datetime.now())
    

    if status != True:
        show_alert(f"修改Box ID: {BoxID}數量異常")
        return 
    
    # also renew Inventory data
    Inventory_Info = SQL_handler.fetch_data_from_Inventory_table({"InventoryID":BoxID[:-1]})
    if isinstance(Inventory_Info, str):
        show_alert(f"讀取庫存資料異常: {Inventory_Info}")
        return 
    elif Inventory_Info == []:
        show_alert("資料庫Inventory表內無此資料")
        return
    
    current_total_quantity = Inventory_Info[0][2]

    new_total_quantity = current_total_quantity - int(line_bar3.text())
    status = SQL_handler.update_total_quantity_by_inventory_id(BoxID[:-1], new_total_quantity)

    if status != True:
        show_alert(f"修改Inventory ID: {BoxID[:-1]}總數量異常")
        return 

    
    # if isinstance(back_list, str):
    #     show_alert(back_list)
    # else:
    #     line_bar2.setText(str(back_list[1]))
    #     line_bar_box_num.setText(str(back_list[0]))
    
    
    # execl_handle.export_to_operateion_table(BoxID, "OUT", line_bar3.text(), back_list[0], '')
    if (chk_box_state != 2):
        from longin_page import user
        unicode = SQL_handler.fetch_data_from_Boxes_table({"BoxID":BoxID})[0][4]
        from barcode_generator import placecode, stockroom
        if(get_place_from_code(int(BoxID[17]),placecode) != None):
            place = get_place_from_code(int(BoxID[17]),placecode)
            if(find_stockroom_name(place ,int(BoxID[18]), stockroom) != None):
                source_stcok_room = find_stockroom_name(place ,int(BoxID[18]), stockroom)
        
        status = SQL_handler.add_data_to_manage_table([unicode, "OUT", (line_bar_box_num.text()-new_quantity), source_stcok_room, None, None, None, datetime.now(), user])
        status = SQL_handler.add_data_to_Operation_table([BoxID, "OUT", line_bar_box_num.text(), new_quantity, "", datetime.now(), user])

        if status != True:
            show_alert(f"操作紀錄存入資料庫時發生異常: {status}")
    elif (chk_box_state == 2):
        
        
        today_date = datetime.today().strftime('%Y-%m-%d')
        inventory_id = data1.text()+today_date.replace('-','')[2:]+str(placecode[stockroom[combo2.currentText()][0]])+str(stockroom[combo2.currentText()][1])
        # print(SQL_handler.fetch_data_from_Inventory_table({"InventoryID":inventory_id}))

        inventory_list = [inventory_id, data1.text(),int(line_bar3.text()),combo2.currentText() ,today_date, user]
        # box_num = execl_handle.get_box_num(inventory_id)
        box_num = SQL_handler.get_max_serial_number(inventory_id)
        # print(box_num)
        
        Supplier_batch = Name_Rule_SQL_handler.get_table_data(table_name="boxes", options={"BoxID":BoxID}, database_name="test_pcb") # Test20240829

        # print(Supplier_batch)

        boxes_list = [[inventory_id, int(box_num), int(line_bar3.text()), Supplier_batch[0][5]]]

        # print(boxes_list)
    #     # box_id = boxes_list[i][0]+execl_handle.number_to_letter(int(boxes_list[i][1]))
        in_update(inventory_list, boxes_list, data1, motivation="SWITCH", old_box_id= BoxID)


    from barcode_generator import out_store_page
    out_store_page()

def back_update(line_bar2, line_bar_box_num, data1, data2, data3, data4, line_bar3, date_choose2):
    if (line_bar3.text() == ''):
        show_alert("請填入退庫數量")
        return
    elif (int(line_bar3.text()) > int(line_bar_box_num.text())):
        show_alert("退庫數量不可超過在箱數量")
        return

    from barcode_generator import placecode, stockroom
    # Find back BoxID
    # BoxID = data1.text()+ data2.text().split('/')[0][2:] + data2.text().split('/')[1].zfill(2)+ data2.text().split('/')[2].zfill(2) +str(placecode[stockroom[data3.currentText()][0]])+str(stockroom[data3.currentText()][1])+ execl_handle.number_to_letter(int(data4.text()))
    Back_BoxID = date_choose2.text().split(',')[0]
    BoxID = date_choose2.text().split(',')[1]
    
    # ----------------------------------------------------------------------------------------------------
    SQL_Box = SQL_handler.fetch_data_from_Boxes_table({"BoxID": BoxID})

    if isinstance(SQL_Box, str):
        show_alert(f"讀取庫存資料異常: {SQL_Box}")
        return 
    elif SQL_Box == []:
        show_alert(f"資料庫中沒有BoxID: {BoxID} 資料")
        return
    
    current_quantity = SQL_Box[0][3]

    new_quantity = 0 #current_quantity - int(line_bar3.text())

    status = SQL_handler.update_quantity_by_box_id(BoxID, new_quantity, datetime.now())

    if status != True:
        show_alert(f"更新[{BoxID}]數量時發生異常: {status}")
        return
    
    # also renew Inventory data
    Inventory_Info = SQL_handler.fetch_data_from_Inventory_table({"InventoryID":BoxID[:-1]})
    if isinstance(Inventory_Info, str):
        show_alert(f"讀取庫存資料異常: {Inventory_Info}")
        return 
    elif Inventory_Info == []:
        show_alert("資料庫Inventory表內無此資料")
        return
    
    current_total_quantity = Inventory_Info[0][2]

    new_total_quantity = current_total_quantity - int(line_bar_box_num.text()) #int(line_bar3.text())
    status = SQL_handler.update_total_quantity_by_inventory_id(BoxID[:-1], new_total_quantity)

    if status != True:
        show_alert(f"修改Inventory ID: {BoxID[:-1]}總數量異常")
        return 

    # ----------------------------------------------------------------------------------------------------

    SQL_Box = SQL_handler.fetch_data_from_Boxes_table({"BoxID": Back_BoxID})

    if isinstance(SQL_Box, str):
        show_alert(f"讀取庫存資料異常: {SQL_Box}")
        return 
    elif SQL_Box == []:
        show_alert(f"資料庫中沒有BoxID: {Back_BoxID} 資料")
        return
    
    current_quantity = SQL_Box[0][3]

    new_quantity = current_quantity + int(line_bar3.text())

    status = SQL_handler.update_quantity_by_box_id(Back_BoxID, new_quantity, datetime.now())

    if status != True:
        show_alert(f"更新[{Back_BoxID}]數量時發生異常: {status}")
        return
    

    # also renew Inventory data
    Inventory_Info = SQL_handler.fetch_data_from_Inventory_table({"InventoryID":Back_BoxID[:-1]})
    if isinstance(Inventory_Info, str):
        show_alert(f"讀取庫存資料異常: {Inventory_Info}")
        return 
    elif Inventory_Info == []:
        show_alert("資料庫Inventory表內無此資料")
        return
    
    current_total_quantity = Inventory_Info[0][2]

    new_total_quantity = current_total_quantity + int(line_bar3.text())
    status = SQL_handler.update_total_quantity_by_inventory_id(Back_BoxID[:-1], new_total_quantity)

    if status != True:
        show_alert(f"修改Inventory ID: {Back_BoxID[:-1]}總數量異常")
        return 
    # ----------------------------------------------------------------------------------------------------

    from longin_page import user
    unicode = SQL_handler.fetch_data_from_Boxes_table({"BoxID":BoxID})[0][4]
    from barcode_generator import placecode, stockroom
    if(get_place_from_code(int(BoxID[17]),placecode) != None):
        place = get_place_from_code(int(BoxID[17]),placecode)
        if(find_stockroom_name(place ,int(BoxID[18]), stockroom) != None):
            source_stcok_room = find_stockroom_name(place ,int(BoxID[18]), stockroom)
    if(get_place_from_code(int(Back_BoxID[17]),placecode) != None):
        place = get_place_from_code(int(Back_BoxID[17]),placecode)
        if(find_stockroom_name(place ,int(Back_BoxID[18]), stockroom) != None):
            destination_stcok_room = find_stockroom_name(place ,int(Back_BoxID[18]), stockroom)

    status = SQL_handler.add_data_to_manage_table([unicode, "RETURNED", abs(current_quantity-new_quantity), source_stcok_room, destination_stcok_room, None, None,datetime.now(), user])
    status = SQL_handler.add_data_to_Operation_table([Back_BoxID, "RETURNED", current_quantity, new_quantity, "", datetime.now(), user])

    if status != True:
        show_alert(f"操作紀錄存入資料庫時發生異常: {status}")
        return
    
    show_alert("退庫成功", "通知")

    from barcode_generator import back_store_page
    back_store_page()
    # back_list = execl_handle.back_update_quantity(BoxID, int(line_bar3.text()))
    # if isinstance(back_list, str):
    #     show_alert(back_list)
    # else:
    #     line_bar2.setText(str(back_list[1]))
    #     line_bar_box_num.setText(str(back_list[0]))
    #     line_bar3.setText('')

    # execl_handle.export_to_operateion_table(BoxID, "RETURNED", line_bar3.text(), back_list[0], '')
    
def barcode_reader(line_bar1, line_bar_batch, line_bar2, date_choose1, combo1, combo2, line_bar_box, line_bar_box_num, date_choose2, out_stock=False):
    text, ok= QInputDialog.getText(None, '條碼機', '請使用讀條碼機讀條碼(需切換輸入法為英文)')
    Box_select = SQL_handler.fetch_data_from_Boxes_table({"UniCode":text})
   
    if len(Box_select) == 0:
        # 2 situations: 1. no data in database 2. the barcode is not real infromation, it's aleady switch stock
        Box_select = SQL_handler.fetch_data_from_Boxes_table({"BoxID":text})
        if len(Box_select) == 0:
            show_alert(f"自料庫尚無{text}資訊")
            return
    elif isinstance(Box_select, str):
        show_alert(f"讀取資料時發生錯誤")
        return
    text_now = Box_select[0][0]
    if(ok):
        if(len(text) == 20):
            # line_bar1.setText(text[0:11])
            line_bar1.setText(text_now[0:11])
            # line_bar2.setText(text[19:])
            year = int("20"+text_now[11:13])
            month = int(text_now[13:15])
            day = int(text_now[15:17])
            date_choose1.setDate(QDate(year, month, day))
            from barcode_generator import placecode, stockroom
            if(get_place_from_code(int(text_now[17]),placecode) != None):
                place = get_place_from_code(int(text_now[17]),placecode)
                if(find_stockroom_name(place ,int(text_now[18]), stockroom) != None):
                    stcok_room = find_stockroom_name(place ,int(text_now[18]), stockroom)
                    combo1.setCurrentText(place)
                    combo2.setCurrentText(stcok_room)
                else:
                    show_alert("bar code或QR code有誤!!")

            # line_bar_box.setText(str(execl_handle.letter_to_number(text[-1])))
            # Inventory_select = SQL_handler.fetch_data_from_Inventory_table({"InventoryID":text_now[:-1]})
            # line_bar2.setText(str(Inventory_select[0][2]))
            # Box_select = SQL_handler.fetch_data_from_Boxes_table({"BoxID":text})
            # Box_select = SQL_handler.fetch_data_from_Boxes_table({"UniCode":text})
            line_bar_box_num.setText(str(Box_select[0][3]))
            line_bar_batch.setText(str(Box_select[0][4]))
            batch_total_quantity = 0
            for val in Box_select:
                batch_total_quantity = batch_total_quantity + val[3]
            line_bar2.setText(str(batch_total_quantity))
        else:
            show_alert(f"條碼長度有誤({len(text)})")
            return
        
        # back stock search
        back_box = SQL_handler.fetch_data_from_Operation_table({"BoxID":text_now, "Motion":"SWITCH"})
        # print(back_box)
        if len(back_box) == 0:
            back_box_id = Box_select[0][4]
        elif isinstance(back_box, str):
            show_alert(f"讀取資料時發生錯誤")
            return
        else: 
            back_box_id = back_box[0][5]

        date_choose2.setText(f"{back_box_id},{text_now}")

        if(get_place_from_code(int(back_box_id[17]),placecode) != None):
            place = get_place_from_code(int(back_box_id[17]),placecode)
            if(find_stockroom_name(place ,int(back_box_id[18]), stockroom) != None):
                stcok_room = find_stockroom_name(place ,int(back_box_id[18]), stockroom)
                line_bar_box.setText(f"{place}-{stcok_room}")
            else:
                show_alert("bar code或QR code有誤!!")


def barcode_reader_for_search(text):
    SQL_Box = SQL_handler.fetch_data_from_Boxes_table({"UniCode":text})
    BoxId_Last = SQL_Box[0][0]

    if isinstance(SQL_Box, str):
        show_alert(f"讀取庫存資料異常: {SQL_Box}")
        return False
    elif SQL_Box == []:
        show_alert("無此barcode資料")
        return False

    date = "20"+text[11:13]+'-'+text[13:15]+'-'+text[15:17]

    from barcode_generator import placecode, stockroom
    if(get_place_from_code(int(text[17]),placecode) != None):
        place_org = get_place_from_code(int(text[17]),placecode)
        place = get_place_from_code(int(BoxId_Last[17]),placecode)
        if(find_stockroom_name(place_org ,int(text[18]), stockroom) != None):
            stcok_room_org = find_stockroom_name(place_org ,int(text[18]), stockroom)
        else:
            show_alert("bar code或QR code有誤!!")
        if(find_stockroom_name(place ,int(BoxId_Last[18]), stockroom) != None):
            stcok_room = find_stockroom_name(place ,int(BoxId_Last[18]), stockroom)
        else:
            show_alert("bar code或QR code有誤!!")
    
    quantity = str(SQL_Box[0][3])
    box = str(SQL_Box[0][2])
    
    return [date, stcok_room_org,stcok_room, quantity, box]

def update_Ecount(ERP_Code, input1, datatable, filter):
    status = SQL_handler.update_Ecount_by_ERP_Code(ERP_Code=ERP_Code, Ecount_Code=input1.text())
    if status == True:
        show_alert(f"{ERP_Code}對應Ecount碼已更新成功為{input1.text()}!!", "通知")
        renew_table(datatable, filter)
    else:
        show_alert(f"Error: {status}")

def update_PN(ERP_Code, input1, datatable, filter):
    status = SQL_handler.update_PN_by_ERP_Code(ERP_Code=ERP_Code, PN=input1.text())
    if status == True:
        show_alert(f"{ERP_Code}對應Part Number已更新成功為{input1.text()}!!", "通知")
        renew_table(datatable, filter)
    else:
        show_alert(f"Error: {status}")

def on_button_clicked(b, datatable, filter):
    pyperclip.copy(b)

    material_data = Name_Rule_SQL_handler.get_table_data('material', database_name="pcbmanagement",options={"ERP_Code":b})
    if isinstance(material_data, str):
        show_alert(f"Error: {material_data}")
        return
    
    inventroy_data = Name_Rule_SQL_handler.get_table_data('inventory', database_name="pcbmanagement",options={"ERP_Code":b})
    if isinstance(material_data, str):
        show_alert(f"Error: {material_data}")
        return


    total_count = 0
    for val in inventroy_data:
        total_count = total_count + val[2]


    dialog = QDialog()
    dialog.setWindowTitle(f"物料編碼: {b}")
    dialog.resize(300,300)

    Vlayout = QVBoxLayout()
    Hlayout_1 = QHBoxLayout()
    Hlayout_2 = QHBoxLayout()
    Hlayout_3 = QHBoxLayout()

    label1 = create_label("Ecount碼:",0,0, align= "right",width=100, font_size=18)
    input1 = create_lineedit(0,0,width=100, font_size=10)
    input1.setText(material_data[0][1])
    button_Ecount_update = create_button("更新", "#DDDDDD", 0, 0, height= 30,width=100, font=20)
    button_Ecount_update.clicked.connect(lambda: update_Ecount(b, input1, datatable, filter))
    Hlayout_1.addWidget(label1)
    Hlayout_1.addWidget(input1)
    Hlayout_1.addWidget(button_Ecount_update)
    Vlayout.addLayout(Hlayout_1)

    label3 = create_label("PN:",0,0, align= "right",width=100, font_size=18)
    input3 = create_lineedit(0,0,width=100, font_size=10)
    input3.setText(material_data[0][-2])
    button_PN_update = create_button("更新", "#DDDDDD", 0, 0, height= 30,width=100, font=20)
    button_PN_update.clicked.connect(lambda: update_PN(b, input3, datatable, filter))
    Hlayout_3.addWidget(label3)
    Hlayout_3.addWidget(input3)
    Hlayout_3.addWidget(button_PN_update)
    Vlayout.addLayout(Hlayout_3)


    label2 = create_label("庫存總量:",0,0, align= "right",width=100, font_size=18)
    input2 = create_label("",0,0, align= "left",width=200, font_size=18)
    input2.setText(str(total_count))
    Hlayout_2.addWidget(label2)
    Hlayout_2.addWidget(input2)
    Vlayout.addLayout(Hlayout_2)

    # 添加 QTreeWidget 来显示每个仓库的库存数量
    inventory_tree = QTreeWidget()
    inventory_tree.setHeaderLabels(["倉庫", "數量", "入庫日期"])
    for item in inventroy_data:
        warehouse_item = QTreeWidgetItem([item[3], str(item[2]), item[4].strftime('%Y-%m-%d')])
        inventory_tree.addTopLevelItem(warehouse_item)
    
    Vlayout.addWidget(inventory_tree)
    
    dialog.setLayout(Vlayout)

    dialog.exec_()

def renew_table(datatable, filter):
    datatable.clear()

    header = ["ERP Code", "ECount", "品項名稱", "項目","總類", "尺寸/種類", "%數","容值/阻值/名稱", "電壓", "製造商", "供應商", "PartNumber"]

    datatable.setColumnCount(len(header))
    datatable.setHorizontalHeaderLabels(header)

    SQL_Material = SQL_handler.fetch_data_from_Material_table(filter)

    datatable.setRowCount(len(SQL_Material))

    for idx, data in enumerate(SQL_Material):
        button = QPushButton(SQL_Material[idx][0])
        button.setFixedWidth(80)
        button.clicked.connect(lambda _, b=SQL_Material[idx][0]: on_button_clicked(b, datatable, filter))
        datatable.setCellWidget(idx, 0, button)
        # print(SQL_Material[idx])
        # datatable.setItem(idx, 0, QTableWidgetItem(SQL_Material[idx][0]))
        datatable.setItem(idx, 1, QTableWidgetItem(SQL_Material[idx][1]))
        datatable.setItem(idx, 2, QTableWidgetItem(SQL_Material[idx][2]))
        datatable.setItem(idx, 3, QTableWidgetItem(SQL_Material[idx][3]))
        datatable.setItem(idx, 4, QTableWidgetItem(SQL_Material[idx][4]))
        datatable.setItem(idx, 5, QTableWidgetItem(SQL_Material[idx][5]))
        datatable.setItem(idx, 6, QTableWidgetItem(SQL_Material[idx][6]))
        datatable.setItem(idx, 7, QTableWidgetItem(SQL_Material[idx][7]))
        datatable.setItem(idx, 8, QTableWidgetItem(SQL_Material[idx][8]))
        datatable.setItem(idx, 9, QTableWidgetItem(SQL_Material[idx][9]))
        datatable.setItem(idx, 10, QTableWidgetItem(SQL_Material[idx][10]))
        datatable.setItem(idx, 11, QTableWidgetItem(SQL_Material[idx][12]))

    datatable.resizeColumnsToContents()


def search_material(filter):
    dialog = QDialog()
    dialog.setWindowTitle("搜尋結果")
    dialog.resize(1050,550)

    layout = QVBoxLayout()

    datatable = create_table()

    renew_table(datatable, filter)

    layout.addWidget(datatable)

    dialog.setLayout(layout)
        
    dialog.exec_()


