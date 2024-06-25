from PyQt5.QtWidgets import QWidget, QGridLayout, QScrollArea, QLabel, QPushButton, QLineEdit, QComboBox, QDateEdit, QDialog, QVBoxLayout, QGraphicsDropShadowEffect, QInputDialog
from PyQt5.QtGui import QCursor, QColor, QPixmap
from PyQt5 import QtCore
from PyQt5.QtCore import Qt, QDate
from printer import tag_printer
from tag_pic_creater import draw_tag_sticker
import execl_handle
import os

grid = QGridLayout()

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
    "button":[], #part code input
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

barcode_page_widgets={
    "button1":[],
    "button2":[],
    "button3":[],
    "button4":[],
    "button_sure":[],
    "button_input":[],
    "label1":[],
    "label2":[],
    "label3":[],
    "label4":[],
    "label5":[],
    "label6":[],
    "label7":[],
    "label8":[],
    "line_bar1":[],
    "line_bar2":[],
    "line_bar3":[],
    "date_choose1":[],
    "date_choose2":[],
    "combo1":[],
    "combo2":[],
    "combo3":[],
}

main_page_widgets = {
    "title":[], 
    "button1":[],
    "button2":[], 
    "button3":[],
    "button4":[],
}

def clear_widgets(widgets):
    ''' hide all existing widgets and erase
        them from the global dictionary'''
    for widget in widgets:
        if widgets[widget] != []:
            widgets[widget][-1].hide()
        for i in range(0, len(widgets[widget])):
            widgets[widget].pop()

def create_label(name, l_margin, r_margin, align = 'left'):
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
    if(align == 'center'):
        label.setAlignment(Qt.AlignVCenter | Qt.AlignHCenter)
    elif(align == 'right'):
        label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
    return label

def create_lineedit(l_margin, r_margin, width = 300):
    lineedit = QLineEdit()
    lineedit.setFixedWidth(width)
    lineedit.setFixedHeight(44)
    lineedit.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        '''
            background: '#D9D9D9';
            font-size: 20px;
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

def create_dateedit(l_margin, r_margin, width = 300):
    dateedit = QDateEdit()
    dateedit.setFixedWidth(width)
    dateedit.setFixedHeight(44)
    dateedit.setStyleSheet(
        #setting variable margins
        "*{margin-left: " + str(l_margin) +"px;"+
        "margin-right: " + str(r_margin) +"px;"+
        '''
            background: '#D9D9D9';
            font-size: 20px;
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

def create_button(name, color ,l_margin, r_margin, height = 44, width = 110, qcolor = [0,0,0, 160]):
    button = QPushButton(name)
    button.setFixedWidth(width)
    button.setFixedHeight(height)
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
    shadow_effect.setColor(QColor(qcolor[0], qcolor[1], qcolor[2], qcolor[3]))
    button.setGraphicsEffect(shadow_effect)
    return button

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

def create_combobox(item_list, l_margin, r_margin, width= 200, font_size = 14):
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
            font-family: Microsoft JhengHei
        }
        *:hover{
            background: '#ECECEC';
        }
        '''
    )
    return combo

def show_alert(message, type="警告"):
    dialog = QDialog()
    dialog.setWindowTitle(type)
    dialog.resize(200,100)
    layout = QVBoxLayout()
    
    layout.addWidget(QLabel(message))

    close_button = QPushButton("Close")
    close_button.clicked.connect(dialog.close)
    layout.addWidget(close_button)
    
    dialog.setLayout(layout)
    
    dialog.exec_()

# 定義函數根據part_number返回整列資訊
def get_part_info(df ,part_number):
    part_info = df[df['part_number'] == part_number]
    if not part_info.empty:
        return part_info.values.tolist()
    else:
        return None

def close(dialog, num):
    for i in range(num):
        os.remove("pic"+str(i)+".png")
    dialog.close()

def tprint(dialog, num, part_number, in_stock_date, count, stock_place):
    execl_handle.execl_stock_record(part_number, in_stock_date, count, stock_place)

    for i in range(num):
        tag_printer("pic"+str(i)+".png")
        os.remove("pic"+str(i)+".png")
    dialog.close()

def preview(num, part_code, count, in_stock_time, create_time, stock_place , total_num, check_code: int):
    if(check_code == 0):
        if(total_num != count*num):
            show_alert("數量加總與總數量不同!")
            return
    elif(check_code == 1):
        count = total_num - count
    elif(check_code == 2):
        count = total_num - count    

    headers = ["part_number", "品項編號", "品項名稱", "項目", "種類", "尺寸/種類","%數", "容值/阻值/名稱", "電壓", "廠商", "供應商", "產生時間"]
    file_path = 'output.xlsx'
    
    if(len(part_code) != 11):
        show_alert("輸入編碼長度錯誤!")
        return
    
    output_df = execl_handle.check_output_existing(file_path, headers)

    if (get_part_info(output_df, part_code) == None):
        show_alert("零件編碼未建立!")
    else:
        _list = get_part_info(output_df, part_code)[-1]

        if (str(_list[4]) == "電容"):
            part_spec = str(_list[3])+", "+str(_list[5])+", "+ str(_list[2]).split(',')[0] +", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
            part_name = "電容/電容"
        elif (str(_list[4]) == "電阻"):
            part_spec = str(_list[3])+", "+str(_list[5])+", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
            part_name = "電阻/電阻"
        else:
            part_spec = str(_list[3])
            part_name = str(_list[4])+"/"+str(_list[5])

        manufacter = _list[9]
        supplier = _list[10]

        dialog = QDialog()
        dialog.setWindowTitle("預覽列印")
        dialog.resize(550,550)
        layout = QVBoxLayout()

        scroll_area = QScrollArea()
        scroll_area_widget_contents = QWidget()
        scroll_layout = QVBoxLayout(scroll_area_widget_contents)

        for i in range(num):
            draw_tag_sticker("pic"+str(i), part_code, part_spec, manufacter, supplier, count, part_name, in_stock_time, create_time, stock_place)
            label = QLabel()
            pixmap = QPixmap("pic"+str(i)+".png")
            label.setPixmap(pixmap)
            scroll_layout.addWidget(label)

        scroll_area.setWidget(scroll_area_widget_contents)
        layout.addWidget(scroll_area)


        close_button = QPushButton("Close")
        close_button.clicked.connect(lambda: close(dialog, num))
        layout.addWidget(close_button)

        print_button = QPushButton("Print")
        print_button.clicked.connect(lambda: tprint(dialog, num, part_code, in_stock_time, count, stock_place))
        layout.addWidget(print_button)
        
        dialog.setLayout(layout)
        
        dialog.exec_()

def find_stockroom_name(location, index, stockroom):
    for key, value in stockroom.items():
        if value[0] == location and value[1] == index:
            return key
    return None

def get_place_from_code(code, placecode):
    reverse_placecode = {v: k for k, v in placecode.items()}
    return reverse_placecode.get(code, None)

def barcode_reader(line_bar1, line_bar2, date_choose1, combo1, combo2):
    text, ok= QInputDialog.getText(None, '條碼機', '請使用讀條碼機讀條碼(需切換輸入法為英文)')
    if(ok):
        if(len(text) > 15):
            line_bar1.setText(text[0:11])
            line_bar2.setText(text[19:])
            year = int("20"+text[11:13])
            month = int(text[13:15])
            day = int(text[15:17])
            date_choose1.setDate(QDate(year, month, day))
            from barcode_generator import placecode, stockroom
            if(get_place_from_code(int(text[17]),placecode) != None):
                place = get_place_from_code(int(text[17]),placecode)
                if(find_stockroom_name(place ,int(text[18]), stockroom) != None):
                    stcok_room = find_stockroom_name(place ,int(text[18]), stockroom)
                    combo1.setCurrentText(place)
                    combo2.setCurrentText(stcok_room)
                else:
                    show_alert("bar code或QR code有誤!!")
