from component import get_place_from_code, find_stockroom_name, get_part_info, grid, barcode_page_widgets as widgets, clear_widgets, main_page_widgets,create_checkBox, create_table, create_button,create_label,create_lineedit, create_dateedit, create_combobox,show_alert, barcode_reader, create_tab, create_textEdit, out_update, back_update, in_update, preview, preview2, create_checkBox_header, preview3
from PyQt5.QtGui import QIntValidator, QColor
from PyQt5.QtWidgets import QGridLayout, QWidget, QTableWidgetItem, QHeaderView, QDialog, QFileDialog, QVBoxLayout, QTableWidget, QHBoxLayout, QSpacerItem, QSizePolicy
from PyQt5.QtCore import QDate, Qt
import execl_handle
import SQL_handler, Name_Rule_SQL_handler
from datetime import datetime
import pandas as pd

stockroom = {"上海成品": ["上海",0],
    "上海成品樣品": ["上海",1],
    "上海零件": ["上海",2],
    "上海零件樣品": ["上海",3],
    "無錫(鈺健)零件": ["無錫",0],
    "無錫(蔣總)零件": ["無錫",1],
    "深圳 - 梁工": ["深圳",0],
    "杭州洲鉅": ["杭州",0],
    "杭州洲鉅 (產線)": ["杭州",1],
    "竹北成品": ["竹北",0],
    "竹北成品樣品": ["竹北",1],
    "竹北成品樣品 - 台南": ["竹北",2],
    "竹北成品樣品 - 維修": ["竹北",3],
    "竹北零件": ["竹北",4],
    "竹北零件樣品": ["竹北",5],
    "竹北組合加工": ["竹北",6],
    "工程小批量預備倉": ["竹北",7],
    "SBIR 專用倉": ["竹北",8],
    "借出倉": ["竹北",9],
    "台南成品 (冰點)": ["台南",0],
    "台南零件倉": ["台南",1],
    "台南半成品倉 (貼片加工)": ["台南",2],
    "台南組合加工倉(產線)": ["台南",3],
    "台南半成品倉( 產線 )": ["台南",4],
    "台南成品倉(產線)": ["台南",5],
    "封樣樣品倉 (台南)": ["台南",6],
    "工具/ 治具倉 (台南)": ["台南",7],
    "客供料倉": ["台南",8],
    "台南報廢倉": ["台南",9],
    "台南費用倉": ["台南","A"],
}

movment_dict = {"入庫": "IN", "出庫": "OUT", "刪除": "DELETE", "轉倉": "SWITCH", "退庫": "RETURNED"} 

stockplace = list(set(value[0] for value in stockroom.values()))
placecode = {"上海":0, "無錫":1, "深圳":2, "杭州":3, "竹北":4, "台南":5}

count = ["1","2","3","4","5","6","7","8","9","10"]

def update_combo2(__stockpalce, combo2):
    keys = [key for key, value in stockroom.items() if value[0] == __stockpalce]
    combo2.clear()
    combo2.addItems(keys) 
    combo2.view().setMinimumWidth(combo2.minimumSizeHint().width()+200)
    combo2.setEnabled(True)

def header_checkbox_clicked(table_print, state):
    for row in range(table_print.rowCount()):
        checkbox = table_print.cellWidget(row, 0)
        if checkbox:
            checkbox.setChecked(state)

def get_in_stock_date_info(df ,date):
    part_info = df[df['入庫日期'] == date]
    if not part_info.empty:
        return part_info.values.tolist()
    else:
        return None

def get_stock_place_info(df ,place):
    part_info = df[df['倉庫'] == place]
    if not part_info.empty:
        return part_info.values.tolist()
    else:
        return None
    
def get_in_stock_count(df,part_number,date,place):
    return df.loc[(df['倉庫'] == place) & (df['入庫日期'] == date) & (df['物料編碼'] == part_number), '總數量'].values[0]

def sure_and_show(part_code,in_stock_date, stock_room, LineEdit, l1, l2, l3, l4, l5, button_print_export,check_code):
    
    if(check_code == 0):
        # headers = ["ERP Code", "品項編號", "品項名稱", "項目", "種類", "尺寸/種類","%數/封裝", "容值/阻值/名稱", "電壓/腳位大小/頻率", "廠商", "供應商", "產生時間", "Part Number"]
        # file_path = 'output.xlsx'

        if(len(part_code) != 11):
            show_alert("數入編碼長度錯誤")
            return

        # output_df = execl_handle.check_output_existing(file_path)
        output_list = SQL_handler.fetch_data_from_Material_table({"ERP_Code": part_code})

        if (output_list == []):
            l1.setVisible(False)
            l2.setVisible(False)
            l3.setVisible(False)
            l4.setVisible(False)
            l5.setVisible(False)
            button_print_export.setVisible(False)
            show_alert("零件編碼未建立")
        else:
            l1.setVisible(True)
            l2.setVisible(True)
            l3.setVisible(True)
            l4.setVisible(True)
            l5.setVisible(True)
            button_print_export.setVisible(True)
    else:
        output_list = SQL_handler.fetch_data_from_Inventory_table({"ERP_Code": part_code, "Location":stock_room, "EntryDate":in_stock_date})
        # file_path = 'stock.xlsx'
        # output_df = execl_handle.check_output_existing(file_path)
        # if (get_part_info(output_df, part_code) != None):
        #     if (get_in_stock_date_info(output_df, in_stock_date) != None):
        #         if (get_stock_place_info(output_df, stock_room) == None):
        #             l1.setVisible(False)
        #             l2.setVisible(False)
        #             l3.setVisible(False)
        #             l4.setVisible(False)
        #             l5.setVisible(False)
        #             button_print_export.setVisible(False)
        #             show_alert("物料編碼["+ part_code +"]無對應倉庫!")
        #         else:
        #             l1.setVisible(True)
        #             l2.setVisible(True)
        #             l3.setVisible(True)
        #             l4.setVisible(True)
        #             l5.setVisible(True)
        #             button_print_export.setVisible(True)
        #             LineEdit.setText(str(get_in_stock_count(output_df, part_code, in_stock_date, stock_room)))
        #     else:
        #         show_alert("物料編碼["+ part_code +"]無對應入庫日期!")
        # else:
        #         show_alert("無對應物料編碼["+ part_code +"]")

        if (output_list == []):
            l1.setVisible(False)
            l2.setVisible(False)
            l3.setVisible(False)
            l4.setVisible(False)
            l5.setVisible(False)
            button_print_export.setVisible(False)
            show_alert("物料編碼["+ part_code +"]無對應資料!")
        else:
            l1.setVisible(True)
            l2.setVisible(True)
            l3.setVisible(True)
            l4.setVisible(True)
            l5.setVisible(True)
            button_print_export.setVisible(True)
            LineEdit.setText(str(output_list[0][2]))

def button1_click():
    in_store_page()

def button2_click():
    out_stock_select_page()

def button3_click():
    back_store_page()

def button_print_click():
    print_page()

def button_operation_search_click():
    operation_search_page()

def barcode_select_page():
    clear_widgets(main_page_widgets)

    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    from main_page import main_page
    button_back.clicked.connect(main_page)

    button1 = create_button("入庫", "#DDDDDD", 0, 0)
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("出庫", "#DDDDDD", 0, 0)
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("退庫", "#DDDDDD", 0, 0)
    widgets["button3"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button_print = create_button("列印/刪除", "#DDDDDD", 0, 0)
    widgets["button_print"].append(button_print)

    #place global widgets on the grid
    grid.addWidget(widgets["button_print"][-1], 1, 4, 1, 1)

    button_operation = create_button("操作查詢", "#DDDDDD", 0, 0)
    widgets["button_operation"].append(button_operation)
    grid.addWidget(widgets["button_operation"][-1], 1, 5, 1, 1)

    button1.clicked.connect(button1_click)
    button2.clicked.connect(button2_click)
    button3.clicked.connect(button3_click)
    button_operation.clicked.connect(button_operation_search_click)
    button_print.clicked.connect(button_print_click)

def switch_stock(state: bool, combo1, combo2):
    combo1.setEnabled(state)
    combo2.setEnabled(state)

def in_stock_inform(table_show, date):
    # read inventory from stock
    header = ["ERP Code", "倉庫", "箱號","目前在箱數量"]
    
    table_show.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

    table_show.setColumnCount(len(header))
    table_show.setHorizontalHeaderLabels(header)

    date_obj = datetime.strptime(date, "%Y/%m/%d")
    formatted_date = date_obj.strftime("%Y-%m-%d")

    SQL_Inventory = SQL_handler.fetch_data_from_Inventory_table({"EntryDate":formatted_date})

    if isinstance(SQL_Inventory, str):
        show_alert(f"讀取庫存資料異常: {SQL_Inventory}")
        return

    box_list = []

    Inventory_IDs = [item[0] for item in SQL_Inventory]
    filters = {"InventoryID": Inventory_IDs}

    SQL_Box = SQL_handler.fetch_data_from_Boxes_table(filters)

    for val in SQL_Box:
        index = next((i for i, item in enumerate(SQL_Inventory) if item[0] == val[1]), None)
        if (index != None):
            box_list.append([SQL_Inventory[index][1], SQL_Inventory[index][3], val])

    table_show.setRowCount(len(box_list))

    for idx, data in enumerate(box_list):
        table_show.setItem(idx, 0, QTableWidgetItem(box_list[idx][0]))
        table_show.setItem(idx, 1, QTableWidgetItem(box_list[idx][1]))
        table_show.setItem(idx, 2, QTableWidgetItem(str(box_list[idx][2][2])))
        table_show.setItem(idx, 3, QTableWidgetItem(str(box_list[idx][2][3])))

def out_stock_inform(table_show, date):
    header = ["ERP Code", "倉庫", "箱號","出庫數量"]

    table_show.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

    table_show.setColumnCount(len(header))
    table_show.setHorizontalHeaderLabels(header)

    date_obj = datetime.strptime(date, "%Y/%m/%d")
    formatted_date = date_obj.strftime("%Y-%m-%d")

    SQL_Inventory = SQL_handler.fetch_data_from_Inventory_table({})

    if isinstance(SQL_Inventory, str):
        show_alert(f"讀取庫存資料異常: {SQL_Inventory}")
        return
    
    box_list = []

    Inventory_IDs = [item[0] for item in SQL_Inventory]
    filters = {"InventoryID": Inventory_IDs}

    SQL_Box = SQL_handler.fetch_data_from_Boxes_table(filters)

    for val in SQL_Box:
        index = next((i for i, item in enumerate(SQL_Inventory) if item[0] == val[1]), None)
        if (index != None):
            box_list.append([SQL_Inventory[index][1], SQL_Inventory[index][3], val])

    operation_list = []

    Box_IDs = [item[2][0] for item in box_list]
    filters = {
        "BoxID": Box_IDs,  # 用多个 BoxID 进行批量查询
        "Motion": "SWITCH",
        "LastUpdated": formatted_date
    }

    SQL_Operation = SQL_handler.fetch_data_from_Operation_table(filters)


    for val in box_list:
        index = next((i for i, item in enumerate(SQL_Operation) if item[1] == val[2][0]), None)
        if(index !=None):
            operation_list.append([val[0], val[1], val[2][2], SQL_Operation[index][4]-SQL_Operation[index][3]])

    table_show.setRowCount(len(operation_list))

    for idx, data in enumerate(operation_list):
        table_show.setItem(idx, 0, QTableWidgetItem(operation_list[idx][0]))
        table_show.setItem(idx, 1, QTableWidgetItem(operation_list[idx][1]))
        table_show.setItem(idx, 2, QTableWidgetItem(str(operation_list[idx][2])))
        table_show.setItem(idx, 3, QTableWidgetItem(str(operation_list[idx][3])))

def search_box_inform(table_print, ERP_Code, date, stock_place, chk_box):

    header = ["", "ERP Code", "入庫日期", "原始倉", "現倉", "箱號","在箱數量", "唯一碼", "供應商批號", "列印次數"]

    table_print.setColumnCount(len(header))
    table_print.setHorizontalHeaderLabels(header)

    filter_dict = {}
    
    if len(ERP_Code) == 20:
        filter_dict.update({"ERP_Code":ERP_Code[0:11]})

    else:

        if ERP_Code != "":
            filter_dict.update({"ERP_Code":ERP_Code})

        if stock_place != "":
            filter_dict.update({"Location":stock_place})

        if date != "":
            filter_dict.update({"EntryDate":date})

    SQL_Inventory = SQL_handler.fetch_data_from_Inventory_table(filter_dict)

    if isinstance(SQL_Inventory, str):
        show_alert(f"讀取庫存資料異常: {SQL_Inventory}")
        return
    
    box_list = []

    # SQL_handler.get_table_data("Boxes")

    Inventory_IDs = [item[0] for item in SQL_Inventory]
    filters = {"InventoryID": Inventory_IDs}

    if len(ERP_Code) == 20:
        SQL_Box = SQL_handler.fetch_data_from_Boxes_table({"BoxID": ERP_Code})
    else:
        SQL_Box = SQL_handler.fetch_data_from_Boxes_table(filters)

    for val in SQL_Box:
        index = next((i for i, item in enumerate(SQL_Inventory) if item[0] == val[1]), None)
        if (index != None):
            formatted_date = SQL_Inventory[index][4].strftime("%Y-%m-%d")
            box_list.append([SQL_Inventory[index][1], formatted_date, SQL_Inventory[index][3], val])
    
    table_print.setRowCount(len(box_list))

    


    chk_box.setVisible(True)

    for idx, data in enumerate(box_list):

        # Get stock by UniCode
        UniCode = box_list[idx][3][4]
        print(UniCode)
        from barcode_generator import placecode, stockroom
        if(get_place_from_code(int(UniCode[17]),placecode) != None):
            place = get_place_from_code(int(UniCode[17]),placecode)
            if(find_stockroom_name(place ,int(UniCode[18]), stockroom) != None):
                stcok_room = find_stockroom_name(place ,int(UniCode[18]), stockroom)
            else:
                show_alert("bar code或QR code有誤!!")

        checkbox = create_checkBox()
        table_print.setCellWidget(idx, 0, checkbox)
        table_print.setItem(idx, 1, QTableWidgetItem(box_list[idx][0])) #ERP Code
        table_print.setItem(idx, 2, QTableWidgetItem(box_list[idx][1])) #入庫日期
        table_print.setItem(idx, 3, QTableWidgetItem(stcok_room)) #原始倉
        table_print.setItem(idx, 4, QTableWidgetItem(box_list[idx][2])) #現倉
        table_print.setItem(idx, 5, QTableWidgetItem(str(box_list[idx][3][2]))) #箱號
        table_print.setItem(idx, 6, QTableWidgetItem(str(box_list[idx][3][3]))) #在箱數量
        table_print.setItem(idx, 7, QTableWidgetItem(str(box_list[idx][3][4]))) #唯一碼
        table_print.setItem(idx, 8, QTableWidgetItem(str(box_list[idx][3][5]))) #供應商批號
        table_print.setItem(idx, 9, QTableWidgetItem(str(box_list[idx][3][7]))) 

    table_print.resizeColumnsToContents()

def box_operation_inform(table_print, ERP_Code, date, stock_place, op_date, movment):

    header = ["唯一碼", "動作", "數量", "原倉庫","目標倉庫", "領料單", "工單","操作時間", "操作者"]

    table_print.setColumnCount(len(header))
    table_print.setHorizontalHeaderLabels(header)
    
    filter_dict = {}

    if ERP_Code != "":
        if len(ERP_Code) == 20:
            Box_data = SQL_handler.fetch_data_from_Boxes_table({"BoxID":ERP_Code})
            filter_dict.update({"UniCode":Box_data[0][4]})
        elif len(ERP_Code) == 11:
            filter_dict.update({"ERPCode":ERP_Code})
        else:
            show_alert("編碼或條碼長度有誤!!")
            return

    # if stock_place != "":
    #     filter_dict.update({"Location":stock_place})

    # if date != "":
    #     filter_dict.update({"EntryDate":date})

    # SQL_Inventory = SQL_handler.fetch_data_from_Inventory_table(filter_dict)

    # if isinstance(SQL_Inventory, str):
    #     show_alert(f"讀取庫存資料異常: {SQL_Inventory}")
    #     return
    
    # box_list = []

    # Inventory_IDs = [item[0] for item in SQL_Inventory]
    # filters = {"InventoryID": Inventory_IDs}

    # SQL_Box = SQL_handler.fetch_data_from_Boxes_table(filters)

    # for val in SQL_Box:
    #     index = next((i for i, item in enumerate(SQL_Inventory) if item[0] == val[1]), None)
    #     if (index != None):
    #         formatted_date = SQL_Inventory[index][4].strftime("%Y-%m-%d")
    #         box_list.append([SQL_Inventory[index][1], formatted_date, SQL_Inventory[index][3], val])

    # for data in SQL_Inventory:
    #     InventoryID = data[0]
    #     SQL_Box = SQL_handler.fetch_data_from_Boxes_table({"InventoryID":InventoryID})
    #     for val in SQL_Box:
    #         # print([data[1], data[3], val])
    #         formatted_date = data[4].strftime("%Y-%m-%d")
    #         box_list.append([data[1], formatted_date, data[3], val])
    
    operation_list = []

    # filters = {}

    # Box_IDs = [item[2][0] for item in box_list]

    if op_date != "":
        filter_dict.update({"UpdateTime":op_date})
    
    if movment != "":
        filter_dict.update({"Motion":movment_dict[movment]})

    SQL_Operation = SQL_handler.fetch_data_from_Manage_table(filter_dict)
    # if (len(SQL_Operation)>= 1):
    #     employee = SQL_handler.fetch_data_from_Employees_table({"EmployeeID":SQL_Operation[0][9]})

    # for val in box_list:
    #     index = next((i for i, item in enumerate(SQL_Operation) if item[1] == val[3][0]), None)
    #     if(index !=None):
    #         formatted_date = SQL_Operation[index][6].strftime("%Y-%m-%d %H:%M:%S")
    #         operation_list.append([val[0], val[1], val[2], val[3][2], SQL_Operation[index][4]-SQL_Operation[index][3], SQL_Operation[index][2], val[3][3], formatted_date, employee[0][1]])


    # for data in box_list:
    #     BoxID = data[3][0]
    #     SQL_Operation = SQL_handler.fetch_data_from_Operation_table({**{"BoxID":BoxID}, **filter_dict})
    #     for val in SQL_Operation:
    #         formatted_date = val[6].strftime("%Y-%m-%d %H:%M:%S")
    #         employee = SQL_handler.fetch_data_from_Employees_table({"EmployeeID":val[7]})

    #         operation_list.append([data[0], data[1], data[2], data[3][2], val[4]-val[3], val[2], data[3][3], formatted_date, employee[0][1]])

    employees = SQL_handler.get_table_data(table_name="employees")

    employees_dict = {item[0]: item[1] for item in employees}

    for data in SQL_Operation:
        formatted_date = data[8].strftime("%Y-%m-%d %H:%M:%S")
        
        operation_list.append([data[1], data[2], data[3], data[4], data[5], data[6], data[7], formatted_date, employees_dict[data[9]]])

    table_print.setRowCount(len(operation_list)) 


    # def sort_by_datetime(item):
    # 返回 datetime 对象
    #     return item[7]

    # sort_operation_list = sorted(operation_list, key = sort_by_datetime)

    for idx, item in enumerate(operation_list):
        table_print.setItem(idx, 0, QTableWidgetItem(item[0]))
        table_print.setItem(idx, 1, QTableWidgetItem(item[1]))
        table_print.setItem(idx, 2, QTableWidgetItem(str(item[2])))
        table_print.setItem(idx, 3, QTableWidgetItem(item[3]))
        table_print.setItem(idx, 4, QTableWidgetItem(item[4]))
        table_print.setItem(idx, 5, QTableWidgetItem(item[5]))
        table_print.setItem(idx, 6, QTableWidgetItem(item[6]))
        table_print.setItem(idx, 7, QTableWidgetItem(item[7]))
        table_print.setItem(idx, 8, QTableWidgetItem(item[8]))

    table_print.resizeColumnsToContents()

def Delete_Box(table_print, ERP_Code, date, stock_place, chk_box):
    from barcode_generator import placecode, stockroom
    checked_items = []
    for row_idx in range(table_print.rowCount()):
        checkbox_widget = table_print.cellWidget(row_idx, 0)
        if checkbox_widget:
            if checkbox_widget.isChecked():
                checked_items.append(row_idx)

    for row_cnt in checked_items:
        Box_ID = table_print.item(row_cnt, 1).text()+table_print.item(row_cnt, 2).text().replace('-','')[2:]+str(placecode[stockroom[table_print.item(row_cnt, 3).text()][0]])+str(stockroom[table_print.item(row_cnt, 3).text()][1])+execl_handle.number_to_letter(int(table_print.item(row_cnt, 5).text()))
        status = SQL_handler.update_quantity_by_box_id(Box_ID, 0)
        if status != True:
            show_alert(f"刪除BoxID: {Box_ID} 時發生錯誤")
        else:
        # quantity_list = execl_handle.out_update_quantity(Box_ID, int(table_print.item(row_cnt, 5).text()))
            # execl_handle.export_to_operateion_table(Box_ID, "DELETE", int(table_print.item(row_cnt, 5).text()), quantity_list[0], '')
            from longin_page import user
            unicode = SQL_handler.fetch_data_from_Boxes_table({"BoxID":Box_ID})[0][4]
            
            if(get_place_from_code(int(Box_ID[17]),placecode) != None):
                place = get_place_from_code(int(Box_ID[17]),placecode)
                if(find_stockroom_name(place ,int(Box_ID[18]), stockroom) != None):
                    source_stcok_room = find_stockroom_name(place ,int(Box_ID[18]), stockroom)
            status = SQL_handler.add_data_to_manage_table([unicode, "DELETE", abs(int(table_print.item(row_cnt, 5).text())), source_stcok_room, None, None, None,datetime.now(), user])
            status = SQL_handler.add_data_to_Operation_table([Box_ID, "DELETE", int(table_print.item(row_cnt, 5).text()), 0, '', datetime.now(), user])

    search_box_inform(table_print, ERP_Code, date, stock_place, chk_box)
    
def printer(table_print):
    checked_items = []
    for row_idx in range(table_print.rowCount()):
        checkbox_widget = table_print.cellWidget(row_idx, 0)
        if checkbox_widget:
            if checkbox_widget.isChecked():
                checked_items.append(row_idx)

    from tag_pic_creater import draw_tag_sticker

    # df_material = execl_handle.excel_file_read("output.xlsx","Sheet1")
    # SQL_material = SQL_handler.fetch_data_from_Material_table({})
    
    today = datetime.today().date()
    formatted_today = today.strftime('%Y-%m-%d')

    count_list = []
    code_list = []

    for idx, val in enumerate(checked_items):
        pic_name = f"sticker/pic{str(idx)}"
        part_code = table_print.item(val, 1).text()
        # _list = get_part_info(df_material, part_code)[-1]
        data = SQL_handler.fetch_data_from_Material_table({"ERP_Code":part_code})
        if data == []:
           show_alert("資料庫查無物料編碼: {part_code}")
        else:
            _list = data[0]
            
            if (str(_list[4]) == "電容"):
                part_spec = str(_list[3])+", "+str(_list[5]).zfill(4)+", "+ str(_list[2]).split(',')[0] +", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
                part_type = "電容/電容"
            elif (str(_list[4]) == "電阻"):
                part_spec = str(_list[3])+", "+str(_list[5]).zfill(4)+", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
                part_type = "電阻/電阻"
            elif ("電容" in str(_list[4])):
                part_spec = str(_list[3])+", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
                part_type = str(_list[4])+"/"+str(_list[5])
            elif ("機構元件" in str(_list[4]) or "晶振" in str(_list[4])):
                part_spec = str(_list[3])+", "+str(_list[6])+", "+str(_list[7])+", "+str(_list[8])
                part_type = str(_list[4])+"/"+str(_list[5])
            elif("PCB" in str(_list[4])):
                part_spec = str(_list[3])+"."+str(_list[8])+", "+str(_list[7])
                part_type = str(_list[4])+"/"+str(_list[5])
            else:
                part_spec = str(_list[3])+", "+str(_list[7]).split('(')[0].split('/')[0]
                part_type = str(_list[4])+"/"+str(_list[5])
            count = int(table_print.item(val, 6).text())
            part_manufacture = _list[9]
            part_supplier = _list[10]
            pn = str(_list[12])
            in_stock_date = table_print.item(val, 2).text()
            box_num = int(table_print.item(val, 5).text())
            create_date = formatted_today
            stockplace = table_print.item(val, 4).text()
            batch = table_print.item(val, 8).text()
            stockplace_org = table_print.item(val, 3).text()

            InventoryID = SQL_handler.fetch_data_from_Inventory_table({"ERP_Code":part_code, "Location":stockplace, "EntryDate":in_stock_date})
            
            draw_tag_sticker(pic_name=pic_name, part_code= part_code, part_spec= part_spec, part_manufacture= part_manufacture, part_supplier=part_supplier, count= count, part_type= part_type, in_stock_date= in_stock_date, create_date= create_date, stockplace= stockplace, pn=pn,box_num=box_num, batch_number=batch, stcok_org = stockplace_org)
            code_list.append(table_print.item(val, 1).text())
            count_list.append([InventoryID[0][0], box_num, count])

    preview2(len(checked_items), code_list, count_list)

def combo_upate_tab(tab_inside, number, tabs, text_edit, total, label_remand):
    tabs.clear()
    text_edit.clear()

    grid.addWidget(widgets['Tab'][-1],3,3,2,2)

    for i in range(number):
        locals()['tab'+str(i)] = QWidget()
        tabs.addTab(locals()['tab'+str(i)],str(i+1))
        layout = QGridLayout()
        layout.setSpacing(10)
        locals()['linedit'+str(i)] = create_lineedit(0, 0, width=80)
        tab_inside["box"+str(i)].append(locals()['linedit'+str(i)])
        label = create_label("數量", 0, 0, 'right')
        layout.addWidget(label,1,1,1,1)
        layout.addWidget(tab_inside["box"+str(i)][-1],1,2,1,2)
        # layout.addRow("數量", tab_inside["box"+str(i)][-1])
        locals()['tab'+str(i)].setLayout(layout)
        # 批號 20240820
        locals()['batch_linedit'+str(i)] = create_lineedit(0, 0, width=150)
        tab_inside["batch_box"+str(i)].append(locals()['batch_linedit'+str(i)])
        label = create_label("供應商批號", 0, 0, 'right')
        layout.addWidget(label,2,1,1,1)
        layout.addWidget(tab_inside["batch_box"+str(i)][-1],2,2,1,2)

        validator = QIntValidator()
        tab_inside["box"+str(i)][-1].setValidator(validator)
        text_edit.append("第"+str(i+1)+"筆數量: ")
        tab_inside["box"+str(i)][-1].editingFinished.connect(lambda:box_edit(text_edit, tab_inside, number, total, label_remand))
        tab_inside["batch_box"+str(i)][-1].editingFinished.connect(lambda:box_edit(text_edit, tab_inside, number, total, label_remand))

    if(number == 1):
        tab_inside["box0"][-1].setText(str(total))

def date_change(lineedit_cover, date_choose1):
    date = date_choose1.date()
    formatted_date = date.toString("yyyy-MM-dd")
    lineedit_cover.setText(formatted_date)

def set_today(lineedit_cover):

    current_date = datetime.now().date()
    formatted_date = current_date.strftime("%Y-%m-%d")
    lineedit_cover.setText(formatted_date)

def box_edit(text_edit, tab_inside, number, total_quantity, label_remand):

    text_edit.clear()
    quantity = 0
    for i in range(number):
        text_edit.append(f"第{str(i+1)}筆數量: {tab_inside['box'+str(i)][-1].text()}, 批號: {tab_inside['batch_box'+str(i)][-1].text()}")
        if (tab_inside['box'+str(i)][-1].text() != ''):
            quantity += int(tab_inside['box'+str(i)][-1].text())

    label_remand.setText("餘: "+str(int(total_quantity) - quantity))
        # print(target_text)
        # new_text = tab_inside['box'+str(i)][-1].text()
       
        # text_edit.textCursor().movePosition(QTextCursor.Right, QTextCursor.MoveAnchor, len(target_text))

        # # 删除现有的数值
        # text_edit.textCursor().select(QTextCursor.WordUnderCursor)
        # text_edit.textCursor().removeSelectedText()

        # text_edit.textCursor().insertText(new_text)     

def export_to_excel():
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    file_path, _ = QFileDialog.getSaveFileName(None, "Save File", "stock.xlsx", "Excel Files (*.xlsx);", options=options)
    if file_path:
        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'
        
        data1 = SQL_handler.get_table_data('Inventory')
        data2 = SQL_handler.get_table_data('Boxes')
        data3 = SQL_handler.get_table_data('Manage')
        empolyee = SQL_handler.get_table_data('Employees')

        # data3 = [t[:5] + t[6:] for t in data3]

        def replace_value(data):
            renew_list = []
            for t in data:
                temp = list(t)
                for i, item in enumerate(temp):
                    if i == 2:
                        if item == "IN":
                            temp[i] = "入庫"
                        elif item == "OUT":
                            temp[i] = "出庫"
                        elif item == "DELETE":
                            temp[i] = "刪除"
                        elif item == "RETURNED":
                            temp[i] = "退庫"
                        elif item == "SWITCH":
                            temp[i] = "轉倉"
                    elif i == 9:
                        temp[i] = empolyee[temp[i]-1][1]
                renew_list.append(tuple(temp))
            return renew_list

        def replace_date(data):
            new_list = [
                t[:4] + (t[4].date(),) + t[5:]  
                for t in data
            ]
            return new_list
        
        data1 = replace_date(data1)
        data3 = replace_value(data3)

        header1 = ['InventoryID', 'ERP_code', 'TotalQuantity', 'Location', 'EntryDate', 'Editby']
        header2 = ['BoxID', 'InventoryID', 'Number', 'Quantity', 'UniCode', 'Supplier_Batch', 'In_Time','Print_Count']
        header3 = ['編號', '唯一碼', '動作', '數量', '來源倉庫', '目的倉庫', '領料單', '工單', '更新時間', '操作者']

        # rows = table.rowCount()
        # columns = table.columnCount()
        # data = []

        # for row in range(rows):
        #     row_data = []
        #     for column in range(columns):
        #         item = table.item(row, column)
        #         row_data.append(item.text() if item else '')
        #     data.append(row_data)

        # header = [table.horizontalHeaderItem(i).text() for i in range(columns)]
        
        # # Extract data from the second table
        # rows2 = table2.rowCount()
        # columns2 = table2.columnCount()

        # data2 = []
        # for row in range(rows2):
        #     row_data = []
        #     for column in range(columns2):
        #         item = table2.item(row, column)
        #         row_data.append(item.text() if item else '')
        #     data2.append(row_data)

        # header2 = [table2.horizontalHeaderItem(i).text() for i in range(columns2)]

        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df1 = pd.DataFrame(data1, columns=header1)
                df2 = pd.DataFrame(data2, columns=header2)
                df3 = pd.DataFrame(data3, columns=header3)
                
                df1.to_excel(writer, sheet_name='Inventory', index=False)
                df2.to_excel(writer, sheet_name='Boxes', index=False)
                df3.to_excel(writer, sheet_name='Operation', index=False)
            
            show_alert(f"儲存{file_path}成功!!", "通知")
        except Exception as e:
            show_alert(f"儲存至{file_path}時發生錯誤: {str(e)}")

def in_store_page():
    clear_widgets(widgets)

    tab_inside = {
        "box0":[],
        "box1":[],
        "box2":[],
        "box3":[],
        "box4":[],
        "box5":[],
        "box6":[],
        "box7":[],
        "box8":[],
        "box9":[],
        "batch_box0":[],
        "batch_box1":[],
        "batch_box2":[],
        "batch_box3":[],
        "batch_box4":[],
        "batch_box5":[],
        "batch_box6":[],
        "batch_box7":[],
        "batch_box8":[],
        "batch_box9":[],
    }

    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    from main_page import main_page
    button_back.clicked.connect(main_page)

    button1 = create_button("入庫", "#DEF100", 0, 0, qcolor=[157,170,0, 60])
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("出庫", "#DDDDDD", 0, 0)
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("退庫", "#DDDDDD", 0, 0)
    widgets["button3"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button_print = create_button("列印/刪除", "#DDDDDD", 0, 0)
    widgets["button_print"].append(button_print)

    #place global widgets on the grid
    grid.addWidget(widgets["button_print"][-1], 1, 4, 1, 1)

    button_operation = create_button("操作查詢", "#DDDDDD", 0, 0)
    widgets["button_operation"].append(button_operation)

    #place global widgets on the grid
    grid.addWidget(widgets["button_operation"][-1], 1, 5, 1, 1)

    label1 = create_label("編號", 0, 0, align='center')
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 2, 1, 1, 1)

    #LineEdit widget
    line_bar1 = create_lineedit(0,0,width=170)
    widgets["line_bar1"].append(line_bar1)
    grid.addWidget(widgets["line_bar1"][-1], 2, 2, 1, 2)

    label2 = create_label("總數量", 0, 0, align='center')
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 3, 1, 1, 1)

    #LineEdit widget
    line_bar2 = create_lineedit(0,0,width=170)
    widgets["line_bar2"].append(line_bar2)
    grid.addWidget(widgets["line_bar2"][-1], 3, 2, 1, 2)

    label3 = create_label("入庫日期", 0, 0, align='center')
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 4, 1, 1, 1)

    date_choose1 = create_dateedit(0,0,width = 170)
    widgets["date_choose1"].append(date_choose1)
    grid.addWidget(widgets["date_choose1"][-1], 4, 2, 1, 2)
    # date_choose1.setEnabled(False)

    # label4 = create_label("製表日期", 0, 0, align='center')
    # widgets["label4"].append(label4)
    # grid.addWidget(widgets["label4"][-1], 5, 1, 1, 1)

    # date_choose2 = create_dateedit(0,0,width = 170)
    # widgets["date_choose2"].append(date_choose2)
    # grid.addWidget(widgets["date_choose2"][-1], 5, 2, 1, 2)

    label5 = create_label("樣品倉地區", 0, 0, align='center')
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 5, 1, 1, 1)

    combo1 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo1"].append(combo1)
    grid.addWidget(widgets["combo1"][-1], 5, 2, 1, 2)

    label6 = create_label("樣品倉", 0, 0, align='center')
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 6, 1, 1, 1)

    combo2 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo2"].append(combo2)
    grid.addWidget(widgets["combo2"][-1], 6, 2, 1, 2)

    button_sure = create_button("資料確認", "#008E8E", 0 ,0)
    widgets["button_sure"].append(button_sure)
    grid.addWidget(widgets["button_sure"][-1], 8, 1, 1, 1)

    label7 = create_label("入庫箱(捲)數", 0, 0, align='right')
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 2, 3, 1, 1)
    label7.setVisible(False)

    combo3 = create_combobox(count, 0, 0, width=140, font_size=20)
    widgets["combo3"].append(combo3)
    grid.addWidget(widgets["combo3"][-1], 2, 4, 1, 2)
    combo3.setEditable(True)
    combo3.setVisible(False)     

    # label8 = create_label("每筆數量", 0, 0, align='right')
    # widgets["label8"].append(label8)
    # grid.addWidget(widgets["label8"][-1], 3, 4, 1, 1)
    # label8.setVisible(False)

    #LineEdit widget
    # line_bar3 = create_lineedit(0, 0, width=140)
    # widgets["line_bar3"].append(line_bar3)
    # grid.addWidget(widgets["line_bar3"][-1], 3, 6, 1, 1)
    # line_bar3.setVisible(False)

    tabs = create_tab(50,0,20,15)
    widgets['Tab'].append(tabs)
    grid.addWidget(widgets['Tab'][-1],3,3,2,2)
    tabs.setVisible(False)

    label_remand = create_label("",0,0,'right',width=60)
    widgets['label_remand'].append(label_remand)
    grid.addWidget(widgets['label_remand'][-1],3,4,1,1,alignment=Qt.AlignRight)

    text_edit = create_textEdit(0,0,20)
    widgets["text_edit"].append(text_edit)
    #place global widgets on the grid
    grid.addWidget(widgets["text_edit"][-1], 5, 3, 3, 2)
    text_edit.setVisible(False)
    
    button4 = create_button("入檔", "#DDDDDD", 0, 0)
    widgets["button4"].append(button4)
    button4.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["button4"][-1], 8, 3, 1, 1)

    button_print_export = create_button("入檔+列印", "#DDDDDD", 0, 0, width=150)
    widgets["button_print_export"].append(button_print_export)
    button_print_export.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["button_print_export"][-1], 8, 4, 1, 1)

    table_show = create_table()
    widgets['table_show'].append(table_show)
    grid.addWidget(widgets['table_show'][-1], 3, 5 , 6, 3)

    label_in_date = create_label('入庫日期',0,0,width=60,font_size=15)
    widgets['label_in_date'].append(label_in_date)
    grid.addWidget(widgets['label_in_date'][-1], 2,5,1,1)

    date_in = create_dateedit(0,0,150,15)
    widgets['date_in'].append(date_in)
    grid.addWidget(widgets['date_in'][-1], 2,6,1,1)

    combo2.view().setMinimumWidth(combo2.minimumSizeHint().width())

    validator = QIntValidator()
    line_bar2.setValidator(validator)
    validator = QIntValidator(1, 9)
    combo3.setValidator(validator)
    # line_bar3.setValidator(validator)
    
    in_stock_inform(table_show, date_in.text())

    combo1.currentIndexChanged.connect(lambda: update_combo2(combo1.currentText(),combo2)) # 20240723 OK
    button2.clicked.connect(button2_click) # 20240723 OK
    button3.clicked.connect(button3_click) # 20240723 OK
    button_print.clicked.connect(button_print_click) # 20240723 OK
    date_in.dateChanged.connect(lambda: in_stock_inform(table_show, date_in.text())) # 20240723 OK
    widgets["combo3"][-1].currentIndexChanged.connect(lambda: combo_upate_tab(tab_inside, int(widgets["combo3"][-1].currentText()), tabs, text_edit, widgets["line_bar2"][-1].text(),  widgets["label_remand"][-1])) # 20240723 OK
    
    button_sure.clicked.connect(lambda: sure_and_show(line_bar1.text(), date_choose1.date().toString("yyyy-MM-dd"),combo2.currentText(),line_bar2 , label7, tabs, combo3, text_edit, button4, button_print_export, 0)) # 20240723 OK
    button4.clicked.connect(lambda: preview(combo3.currentText(), line_bar1.text(), tab_inside, date_choose1.date().toString("yyyy-MM-dd"), combo2.currentText(), line_bar2, 0, line_bar1, table_show, date_in))
    button_print_export.clicked.connect(lambda: preview3(combo3.currentText(), line_bar1.text(), tab_inside, date_choose1.date().toString("yyyy-MM-dd"), combo2.currentText(), line_bar2, 0, line_bar1, table_show, date_in)) # 20240723 OK
    button_operation.clicked.connect(button_operation_search_click) # 20240723 OK

def out_store_page():
    clear_widgets(widgets)

    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    from main_page import main_page
    button_back.clicked.connect(main_page)

    button1 = create_button("入庫", "#DDDDDD", 0, 0)
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("出庫", "#DEF100", 0, 0, qcolor=[157,170,0, 60])
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("退庫", "#DDDDDD", 0, 0)
    widgets["button3"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button_print = create_button("列印/刪除", "#DDDDDD", 0, 0)
    widgets["button_print"].append(button_print)

    #place global widgets on the grid
    grid.addWidget(widgets["button_print"][-1], 1, 4, 1, 1)

    button_operation = create_button("操作查詢", "#DDDDDD", 0, 0)
    widgets["button_operation"].append(button_operation)

    #place global widgets on the grid
    grid.addWidget(widgets["button_operation"][-1], 1, 5, 1, 1)

    # label1 = create_label("料號", 0, 0, align='center')
    # widgets["label1"].append(label1)
    # grid.addWidget(widgets["label1"][-1], 2, 1, 1, 1)

    # button_batch_out = create_button("領料單出庫", "#DDDDDD",0,0)
    # widgets["button_batch_out"].append(button_batch_out)
    # grid.addWidget(widgets["button_batch_out"][-1], 2, 1, 1, 1)
    # button_batch_out.clicked.connect(batch_out_store_page)

    button_input = create_button("條碼機", "#DDDDDD", 0 ,0)
    widgets["button_input"].append(button_input)
    grid.addWidget(widgets["button_input"][-1], 3, 1, 1, 1)

    # line_bar_hide = create_lineedit(0,0,width=170)
    # widgets["line_bar_hide"].append(line_bar_hide)
    # grid.addWidget(widgets["line_bar_hide"][-1], 2, 2, 1, 2)

    #LineEdit widget
    line_bar1 = create_lineedit(0,0,width=170)
    widgets["line_bar1"].append(line_bar1)
    grid.addWidget(widgets["line_bar1"][-1], 3, 2, 1, 2)
    line_bar1.setReadOnly(True)

    label9 = create_label("批號", 0, 0, align='center')
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 4, 1, 1, 1)

    line_bar_batch = create_lineedit(0,0,width=220, font_size=18)
    widgets["line_bar_batch"].append(line_bar_batch)
    grid.addWidget(widgets["line_bar_batch"][-1],4, 2, 1, 2)
    line_bar_batch.setReadOnly(True) 

    label2 = create_label("批號總數量", 0, 0, align='center')
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 5, 1, 1, 1)

    #LineEdit widget
    line_bar2 = create_lineedit(0,0,width=170)
    widgets["line_bar2"].append(line_bar2)
    grid.addWidget(widgets["line_bar2"][-1], 5, 2, 1, 2)
    line_bar2.setReadOnly(True)

    label3 = create_label("入庫日期", 0, 0, align='center')
    widgets["label3"].append(label3)
    # grid.addWidget(widgets["label3"][-1], 5, 1, 1, 1)

    date_choose1 = create_dateedit(0,0,width = 170)
    widgets["date_choose1"].append(date_choose1)
    # grid.addWidget(widgets["date_choose1"][-1], 5, 2, 1, 2)
    date_choose1.setReadOnly(True)

    # label4 = create_label("製表日期", 0, 0, align='center')
    # widgets["label4"].append(label4)
    # grid.addWidget(widgets["label4"][-1], 5, 1, 1, 1)

    date_choose2 = create_lineedit(0,0,width = 170)
    widgets["date_choose2"].append(date_choose2)
    # grid.addWidget(widgets["date_choose2"][-1], 5, 2, 1, 2)

    label5 = create_label("現倉地區", 0, 0, align='center')
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 6, 1, 1, 1)

    combo1 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo1"].append(combo1)
    grid.addWidget(widgets["combo1"][-1], 6, 2, 1, 2)
    combo1.setDisabled(True)

    label6 = create_label("現倉", 0, 0, align='center')
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 7, 1, 1, 1)
    
    combo2 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo2"].append(combo2)
    grid.addWidget(widgets["combo2"][-1], 7, 2, 1, 2)
    combo2.setDisabled(True)

    label_box = create_label("箱號", 0, 0, align='center')
    widgets["label_box"].append(label_box)
    # grid.addWidget(widgets["label_box"][-1], 8, 1, 1, 1)

    line_bar_box = create_lineedit(0, 0, width=170)
    widgets["line_bar_box"].append(line_bar_box)
    # grid.addWidget(widgets["line_bar_box"][-1], 8, 2, 1, 2)
    line_bar_box.setReadOnly(True)

    # button_sure = create_button("資料確認", "#008E8E", 0 ,0)
    # widgets["button_sure"].append(button_sure)
    # grid.addWidget(widgets["button_sure"][-1], 8, 1, 1, 1)

    # label7 = create_label("列印張數", 0, 0, align='right')
    # widgets["label7"].append(label7)
    # grid.addWidget(widgets["label7"][-1], 2, 4, 1, 1)
    # label7.setVisible(False)

    # combo3 = create_combobox(count, 0, 0, width=140, font_size=20)
    # combo3.setCurrentIndex(0)
    # widgets["combo3"].append(combo3)
    # grid.addWidget(widgets["combo3"][-1], 2, 6, 1, 1)
    # combo3.setEditable(False)
    # combo3.setDisabled(True)
    # combo3.setVisible(False)

    label_box_num = create_label("在箱數量", 0, 0, align='right')
    widgets["label_box_num"].append(label_box_num)
    grid.addWidget(widgets["label_box_num"][-1], 3, 3, 1, 1)

    line_bar_box_num = create_lineedit(0, 0, width=140)
    widgets["line_bar_box_num"].append(line_bar_box_num)
    grid.addWidget(widgets["line_bar_box_num"][-1], 3, 4, 1, 1)
    line_bar_box_num.setReadOnly(True)

    label8 = create_label("出庫數量", 0, 0, align='right')
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 4, 3, 1, 1)
    # label8.setVisible(False)

    #LineEdit widget
    line_bar3 = create_lineedit(0, 0, width=140)
    widgets["line_bar3"].append(line_bar3)
    grid.addWidget(widgets["line_bar3"][-1], 4, 4, 1, 1)
    # line_bar3.setVisible(False)

    button4 = create_button("確認出庫", "#DDDDDD", 0, 0)
    widgets["button4"].append(button4)
    # button4.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["button4"][-1], 8, 4, 1, 1)

    chk_box = create_checkBox("是否轉倉")
    widgets["chk_box"].append(chk_box)
    chk_box.setChecked(True)
    chk_box.setEnabled(False)

    grid.addWidget(widgets["chk_box"][-1], 5, 3, 1, 2)

    label_rollover_area = create_label("地區(轉倉)",0,0,align='right')
    widgets["label_rollover_area"].append(label_rollover_area)

    grid.addWidget(widgets["label_rollover_area"][-1], 6, 3, 1, 1)

    label_rollover_place = create_label("倉(轉倉)",0,0,align='right')
    widgets["label_rollover_place"].append(label_rollover_place)

    grid.addWidget(widgets["label_rollover_place"][-1], 7, 3, 1, 1)

    combo_rollover_area = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo_rollover_area"].append(combo_rollover_area)
    grid.addWidget(widgets["combo_rollover_area"][-1], 6, 4, 1, 2)
    # combo_rollover_area.setEnabled(False)

    combo_rollover_place = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo_rollover_place"].append(combo_rollover_place)
    grid.addWidget(widgets["combo_rollover_place"][-1], 7, 4, 1, 2)
    combo_rollover_place.setEnabled(False)

    table_show = create_table()
    widgets['table_show'].append(table_show)
    grid.addWidget(widgets['table_show'][-1], 3, 5 , 6, 3)

    label_in_date = create_label('入庫日期',0,0,width=60,font_size=15)
    widgets['label_in_date'].append(label_in_date)
    grid.addWidget(widgets['label_in_date'][-1], 2,5,1,1)

    date_in = create_dateedit(0,0,150,15)
    widgets['date_in'].append(date_in)
    grid.addWidget(widgets['date_in'][-1], 2,6,1,1)

    combo2.view().setMinimumWidth(combo2.minimumSizeHint().width()+20)
    combo_rollover_place.view().setMinimumWidth(combo_rollover_place.minimumSizeHint().width()+20)

    validator = QIntValidator()
    # combo3.setValidator(validator)
    line_bar3.setValidator(validator)

    out_stock_inform(table_show, date_in.text())

    combo1.currentIndexChanged.connect(lambda: update_combo2(combo1.currentText(),combo2)) # 20240723 OK
    button1.clicked.connect(button1_click) # 20240723 OK
    button3.clicked.connect(button3_click) # 20240723 OK
    button_print.clicked.connect(button_print_click) # 20240723 OK
    # button_sure.clicked.connect(lambda: sure_and_show(line_bar1.text(), date_choose1.date().toString("yyyy-MM-dd"),combo2.currentText(),line_bar2, label7, label8, combo3, line_bar3, button4, 1))
    button_input.clicked.connect(lambda: barcode_reader(line_bar1, line_bar_batch, line_bar2, date_choose1, combo1, combo2, line_bar_box, line_bar_box_num, date_choose2)) # 20240723 OK
    button4.clicked.connect(lambda: out_update(line_bar2, line_bar_box_num, line_bar1, date_choose1, combo2 , line_bar_box, line_bar3, chk_box.checkState(), combo_rollover_area, combo_rollover_place, date_choose2)) # 20240724 OK
    button_operation.clicked.connect(button_operation_search_click) # 20240723 OK
    combo_rollover_area.currentIndexChanged.connect(lambda: update_combo2(combo_rollover_area.currentText(),combo_rollover_place)) # 20240723 OK
    chk_box.stateChanged.connect(lambda:switch_stock(chk_box.checkState(), combo_rollover_area, combo_rollover_place)) # 20240723 OK
    date_in.dateChanged.connect(lambda: out_stock_inform(table_show, date_in.text())) # 20240724 OK

def batch_out_store_page():
    clear_widgets(widgets)

    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    from main_page import main_page
    button_back.clicked.connect(main_page)

    button1 = create_button("入庫", "#DDDDDD", 0, 0)
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("出庫", "#DEF100", 0, 0, qcolor=[157,170,0, 60])
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("退庫", "#DDDDDD", 0, 0)
    widgets["button3"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button_print = create_button("列印/刪除", "#DDDDDD", 0, 0)
    widgets["button_print"].append(button_print)

    #place global widgets on the grid
    grid.addWidget(widgets["button_print"][-1], 1, 4, 1, 1)

    button_operation = create_button("操作查詢", "#DDDDDD", 0, 0)
    widgets["button_operation"].append(button_operation)

    #place global widgets on the grid
    grid.addWidget(widgets["button_operation"][-1], 1, 5, 1, 1)

    # button_single_out = create_button("單筆出庫", "#DDDDDD",0,0)
    # widgets["button_single_out"].append(button_single_out)
    # grid.addWidget(widgets["button_single_out"][-1], 2, 1, 1, 1)
    # button_single_out.clicked.connect(out_store_page)

    button_input = create_button("讀取領料單", "#DDDDDD", 0, 0)
    widgets["button_input"].append(button_input)
    grid.addWidget(widgets["button_input"][-1], 2, 2, 1, 1)

    button1.clicked.connect(button1_click) # 20240723 OK
    button3.clicked.connect(button3_click) # 20240723 OK
    button_print.clicked.connect(button_print_click) # 20240723 OK
    # button_sure.clicked.connect(lambda: sure_and_show(line_bar1.text(), date_choose1.date().toString("yyyy-MM-dd"),combo2.currentText(),line_bar2, label7, label8, combo3, line_bar3, button4, 1))
    button_operation.clicked.connect(button_operation_search_click) # 20240723 OK
    button_input.clicked.connect(batch_out_store_page)

def back_store_page():
    clear_widgets(widgets)

    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    from main_page import main_page
    button_back.clicked.connect(main_page)

    button1 = create_button("入庫", "#DDDDDD", 0, 0)
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("出庫", "#DDDDDD", 0, 0)
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("退庫", "#DEF100", 0, 0, qcolor=[157,170,0, 60])
    widgets["button3"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button_print = create_button("列印/刪除", "#DDDDDD", 0, 0)
    widgets["button_print"].append(button_print)

    #place global widgets on the grid
    grid.addWidget(widgets["button_print"][-1], 1, 4, 1, 1)

    button_operation = create_button("操作查詢", "#DDDDDD", 0, 0)
    widgets["button_operation"].append(button_operation)

    #place global widgets on the grid
    grid.addWidget(widgets["button_operation"][-1], 1, 5, 1, 1)

    button_input = create_button("條碼機", "#DDDDDD", 0 ,0)
    widgets["button_input"].append(button_input)
    grid.addWidget(widgets["button_input"][-1], 2, 1, 1, 1)

    # line_bar_hide = create_lineedit(0,0,width=170)
    # widgets["line_bar_hide"].append(line_bar_hide)
    # grid.addWidget(widgets["line_bar_hide"][-1], 2, 2, 1, 2)

    #LineEdit widget
    line_bar1 = create_lineedit(0,0,width=170)
    widgets["line_bar1"].append(line_bar1)
    grid.addWidget(widgets["line_bar1"][-1], 2, 2, 1, 2)
    line_bar1.setReadOnly(True)

    label9 = create_label("批號", 0, 0, align='center')
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 3, 1, 1, 1)

    line_bar_batch = create_lineedit(0,0,width=220,font_size=18)
    widgets["line_bar_batch"].append(line_bar_batch)
    grid.addWidget(widgets["line_bar_batch"][-1], 3, 2, 1, 2)
    line_bar_batch.setReadOnly(True)

    label2 = create_label("總在庫數量", 0, 0, align='center')
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 4, 1, 1, 1)

    #LineEdit widget
    line_bar2 = create_lineedit(0,0,width=170)
    widgets["line_bar2"].append(line_bar2)
    grid.addWidget(widgets["line_bar2"][-1], 4, 2, 1, 2)
    line_bar2.setReadOnly(True)

    label3 = create_label("入庫日期", 0, 0, align='center')
    widgets["label3"].append(label3)
    # grid.addWidget(widgets["label3"][-1], 5, 1, 1, 1)

    date_choose1 = create_dateedit(0,0,width = 170)
    widgets["date_choose1"].append(date_choose1)
    # grid.addWidget(widgets["date_choose1"][-1], 5, 2, 1, 2)
    date_choose1.setReadOnly(True)

    # label4 = create_label("製表日期", 0, 0, align='center')
    # widgets["label4"].append(label4)
    # grid.addWidget(widgets["label4"][-1], 5, 1, 1, 1)

    date_choose2 = create_lineedit(0,0,width = 170)
    widgets["date_choose2"].append(date_choose2)
    # grid.addWidget(widgets["date_choose2"][-1], 5, 2, 1, 2)

    label5 = create_label("現倉地區", 0, 0, align='center')
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 5, 1, 1, 1)

    combo1 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo1"].append(combo1)
    grid.addWidget(widgets["combo1"][-1], 5, 2, 1, 2)
    combo1.setDisabled(True)

    label6 = create_label("現倉", 0, 0, align='center')
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 6, 1, 1, 1)

    combo2 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo2"].append(combo2)
    grid.addWidget(widgets["combo2"][-1], 6, 2, 1, 2)
    combo2.setDisabled(True)

    label_box = create_label("返回倉庫", 0, 0, align='center')
    widgets["label_box"].append(label_box)
    grid.addWidget(widgets["label_box"][-1], 7, 1, 1, 1)

    line_bar_box = create_lineedit(0, 0, width=220, font_size=17)
    widgets["line_bar_box"].append(line_bar_box)
    grid.addWidget(widgets["line_bar_box"][-1], 7, 2, 1, 2)
    line_bar_box.setReadOnly(True)

    # button_sure = create_button("資料確認", "#008E8E", 0 ,0)
    # widgets["button_sure"].append(button_sure)
    # grid.addWidget(widgets["button_sure"][-1], 8, 1, 1, 1)

    # label7 = create_label("列印張數", 0, 0, align='right')
    # widgets["label7"].append(label7)
    # grid.addWidget(widgets["label7"][-1], 2, 4, 1, 1)
    # label7.setVisible(False)

    # combo3 = create_combobox(count, 0, 0, width=140, font_size=20)
    # combo3.setCurrentIndex(0)
    # widgets["combo3"].append(combo3)
    # grid.addWidget(widgets["combo3"][-1], 2, 6, 1, 1)
    # combo3.setEditable(False)
    # combo3.setDisabled(True)
    # combo3.setVisible(False)

    label_box_num = create_label("在箱數量", 0, 0, align='right')
    widgets["label_box_num"].append(label_box_num)
    grid.addWidget(widgets["label_box_num"][-1], 2, 4, 1, 1)

    line_bar_box_num = create_lineedit(0, 0, width=140)
    widgets["line_bar_box_num"].append(line_bar_box_num)
    grid.addWidget(widgets["line_bar_box_num"][-1], 2, 6, 1, 1)
    line_bar_box_num.setReadOnly(True)

    label8 = create_label("退庫數量", 0, 0, align='right')
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 3, 4, 1, 1)
    # label8.setVisible(False)

    #LineEdit widget
    line_bar3 = create_lineedit(0, 0, width=140)
    widgets["line_bar3"].append(line_bar3)
    grid.addWidget(widgets["line_bar3"][-1], 3, 6, 1, 1)
    # line_bar3.setVisible(False)

    button4 = create_button("確認退庫", "#DDDDDD", 0, 0)
    widgets["button4"].append(button4)
    # button4.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["button4"][-1], 4, 6, 1, 1)

    

    validator = QIntValidator()
    # combo3.setValidator(validator)
    line_bar3.setValidator(validator)

    combo1.currentIndexChanged.connect(lambda: update_combo2(combo1.currentText(),combo2)) # # 20240725 OK
    button1.clicked.connect(button1_click) # 20240725 OK
    button2.clicked.connect(button2_click) # 20240725 OK
    button_print.clicked.connect(button_print_click) # 20240725 OK
    button_input.clicked.connect(lambda: barcode_reader(line_bar1, line_bar_batch, line_bar2, date_choose1, combo1, combo2, line_bar_box, line_bar_box_num, date_choose2)) # 20240725 OK
    button4.clicked.connect(lambda: back_update(line_bar2, line_bar_box_num, line_bar1, date_choose1, combo2 , line_bar_box, line_bar3, date_choose2)) # 20240725 OK
    button_operation.clicked.connect(button_operation_search_click) # 20240725 OK

def print_page():
    clear_widgets(widgets)

    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    from main_page import main_page
    button_back.clicked.connect(main_page)

    button1 = create_button("入庫", "#DDDDDD", 0, 0)
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("出庫", "#DDDDDD", 0, 0)
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("退庫", "#DDDDDD", 0, 0)
    widgets["button3"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button_print = create_button("列印/刪除", "#DEF100", 0, 0, qcolor=[157,170,0, 60])
    widgets["button_print"].append(button_print)

    #place global widgets on the grid
    grid.addWidget(widgets["button_print"][-1], 1, 4, 1, 1)

    button_operation = create_button("操作查詢", "#DDDDDD", 0, 0)
    widgets["button_operation"].append(button_operation)

    #place global widgets on the grid
    grid.addWidget(widgets["button_operation"][-1], 1, 5, 1, 1)

    table_print = create_table()
    widgets["table_print"].append(table_print)

    grid.addWidget(widgets["table_print"][-1], 2, 2, 7, 6)

    # search
    label1 = create_label("編碼", 0, 0, align='right', width=50)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 2, 0, 1, 1)

    #LineEdit widget
    line_bar1 = create_lineedit(0,0,width=170)
    widgets["line_bar1"].append(line_bar1)
    grid.addWidget(widgets["line_bar1"][-1], 2, 1, 1, 2)

    label2 = create_label("入庫日", 0, 0, align='right', width=60)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 3, 0, 1, 1)

    date_choose1 = create_dateedit(0,0,width = 170)
    widgets["date_choose1"].append(date_choose1)
    grid.addWidget(widgets["date_choose1"][-1], 3, 1, 1, 2)
    # date_choose1.setSpecialValueText(" ")
    # date_choose1.setDate(QDate().fromString("01/01/0001", "dd/MM/yyyy"))

    lineedit_cover = create_lineedit(0,0,width =150)
    widgets['lineedit_cover'].append(lineedit_cover)
    grid.addWidget(widgets["lineedit_cover"][-1], 3, 1, 1, 2)


    label5 = create_label("地區", 0, 0, align='right', width=50)
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 4, 0, 1, 1)


    combo1 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo1"].append(combo1)
    grid.addWidget(widgets["combo1"][-1], 4, 1, 1, 2)

    label6 = create_label("倉庫", 0, 0, align='right', width=50)
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 5, 0, 1, 1)

    combo2 = create_combobox([],0,0,width=170,font_size=20)
    widgets["combo2"].append(combo2)
    grid.addWidget(widgets["combo2"][-1], 5, 1, 1, 2)
    combo2.setEditable(True)

    button4 = create_button("搜尋", "#DDDDDD", 0, 0)
    widgets["button4"].append(button4)
    # button4.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["button4"][-1], 6, 0, 1, 2)

    button_printer = create_button("列印", "#DDDDDD", 0, 0)
    widgets["button_printer"].append(button_printer)
    # button4.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["button_printer"][-1], 7, 0, 1, 2)

    button_delete = create_button("刪除", "#FF8888", 0, 0)
    widgets["button_delete"].append(button_delete)

    grid.addWidget(widgets["button_delete"][-1], 8, 0, 1, 2)


    chk_box = create_checkBox_header()
    widgets['chk_box'].append(chk_box)
    chk_box.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["chk_box"][-1], 1, 2, 2, 1)
    
    set_today(lineedit_cover)

    combo1.currentIndexChanged.connect(lambda: update_combo2(combo1.currentText(),combo2)) # 20240724 OK
    button1.clicked.connect(button1_click) # 20240724 OK
    button2.clicked.connect(button2_click) # 20240724 OK
    button3.clicked.connect(button3_click) # 20240724 OK
    button_print.clicked.connect(button_print_click) # 20240724 OK
    button4.clicked.connect(lambda: search_box_inform(table_print, widgets["line_bar1"][-1].text(), widgets["lineedit_cover"][-1].text(), widgets["combo2"][-1].currentText(), chk_box)) # 20240724 OK
    button_printer.clicked.connect(lambda: printer(table_print)) # 20240724 OK
    chk_box.stateChanged.connect(lambda: header_checkbox_clicked(table_print, chk_box.isChecked())) # 20240724 OK
    date_choose1.dateChanged.connect(lambda: date_change(lineedit_cover, date_choose1)) # 20240724 OK
    button_delete.clicked.connect(lambda: Delete_Box(table_print, widgets["line_bar1"][-1].text(), widgets["lineedit_cover"][-1].text(), widgets["combo2"][-1].currentText(), chk_box)) # 20240724 OK
    button_operation.clicked.connect(button_operation_search_click) # 20240724 OK

def out_stock_select_page():
    clear_widgets(widgets)

    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    from main_page import main_page
    button_back.clicked.connect(main_page)

    button1 = create_button("入庫", "#DDDDDD", 0, 0)
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("出庫", "#DEF100", 0, 0, qcolor=[157,170,0, 60])
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("退庫", "#DDDDDD", 0, 0)
    widgets["button3"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button_print = create_button("列印/刪除", "#DDDDDD", 0, 0)
    widgets["button_print"].append(button_print)

    #place global widgets on the grid
    grid.addWidget(widgets["button_print"][-1], 1, 4, 1, 1)

    button_operation = create_button("操作查詢", "#DDDDDD", 0, 0)
    widgets["button_operation"].append(button_operation)

    #place global widgets on the grid
    grid.addWidget(widgets["button_operation"][-1], 1, 5, 1, 1)


    button_single_out = create_button("單筆出庫", "#DDDDDD",0 ,0, height=150, width= 250)
    widgets["button_single_out"].append(button_single_out)
    button_single_out.clicked.connect(out_store_page)
    grid.addWidget(widgets["button_single_out"][-1], 2, 1, 4, 3)

    button_batch_out = create_button("領料單出庫", "#DDDDDD",0 ,0, height=150, width= 250)
    widgets["button_batch_out"].append(button_batch_out)
    grid.addWidget(widgets["button_batch_out"][-1], 2, 3, 4, 3)
    button_batch_out.clicked.connect(batch_out_store_page)


    button1.clicked.connect(button1_click) # 20240723 OK
    button3.clicked.connect(button3_click) # 20240723 OK
    button_print.clicked.connect(button_print_click) # 20240723 OK
    # button_sure.clicked.connect(lambda: sure_and_show(line_bar1.text(), date_choose1.date().toString("yyyy-MM-dd"),combo2.currentText(),line_bar2, label7, label8, combo3, line_bar3, button4, 1))
    button_operation.clicked.connect(button_operation_search_click) # 20240723 OK

def set_column_background_color(table_widget, column, color):
    for row in range(table_widget.rowCount()):
        item = table_widget.item(row, column)
        if item is not None:
            item.setBackground(QColor(color))
        else:
            # 如果該單元格沒有項目，則創建一個新的 QTableWidgetItem
            new_item = QTableWidgetItem()
            new_item.setBackground(QColor(color))
            table_widget.setItem(row, column, new_item)

def Ecount_compare():
    dialog = QDialog()
    dialog.setWindowTitle(f"Ecount數量對照表")
    dialog.resize(1200,800)

    layout = QVBoxLayout(dialog)

    Hlayout = QHBoxLayout()

    load_button = create_button("讀取Excel檔案","#DDDDDD",0,0,width=200)
    Hlayout.addWidget(load_button)

    result_label = create_label("",0,0,width=600)
    Hlayout.addWidget(result_label)

    spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
    Hlayout.addItem(spacer)

    save_button = create_button("儲存Excel檔案","#DDDDDD",0,0,width=200)
    Hlayout.addWidget(save_button)

    Hlayout.setAlignment(Qt.AlignLeft)

    layout.addLayout(Hlayout)

    # table_widget = create_table()
    table_widget = QTableWidget()
    layout.addWidget(table_widget)

    hide_button = create_button("顯示異常項目","#DDDDDD",0,0,width=200) 
    layout.addWidget(hide_button)

    SQL_stock_dict = dict()
    inventory_data = SQL_handler.get_table_data(table_name="Inventory")

    for item in inventory_data:
        key = item[1]
        quantity = item[2]
        
        if key in SQL_stock_dict:
            SQL_stock_dict[key] += quantity
        else:
            SQL_stock_dict[key] = quantity

    def load_excel():
        # 打開文件選擇對話框，選擇 Excel 檔案
        file_name, _ = QFileDialog.getOpenFileName(dialog, "選擇Excel檔案", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            # 使用 pandas 讀取 Excel 檔案
            df = pd.read_excel(file_name)
            df_filtered = df.iloc[1:len(df)-2, [0, 1, 3, 4]]
            
            # 將資料載入到 QTableWidget 中
            table_widget.setRowCount(df_filtered.shape[0])
            table_widget.setColumnCount(df_filtered.shape[1]+2)
            row_headers = ["Ecount碼", "物料編碼", "規格", "Ecount數量", "資料庫數量", "匹配結果"]
            table_widget.setHorizontalHeaderLabels(row_headers)
            
            discrepancy = 0
            reconciliation = 0
            no_ERP = 0
            not_in_database = 0

            for i in range(0, df_filtered.shape[0]):
                for j in range(df_filtered.shape[1]):
                    table_widget.setItem(i, j, QTableWidgetItem(str(df_filtered.iat[i, j])))
                if (isinstance(df_filtered.iat[i, 1],str) and df_filtered.iat[i, 1] in SQL_stock_dict):
                    table_widget.setItem(i,4, QTableWidgetItem(str(SQL_stock_dict[str(df_filtered.iat[i, 1])])))
                    if (abs(int(table_widget.item(i,3).text()) - int(table_widget.item(i,4).text())) > int(table_widget.item(i,3).text()) * 0.1):
                        # table_widget.item(i,0).setBackground(QColor('yellow'))
                        # table_widget.item(i,1).setBackground(QColor('yellow'))
                        # table_widget.item(i,2).setBackground(QColor('yellow'))
                        # table_widget.item(i,3).setBackground(QColor('yellow'))
                        table_widget.setItem(i,5, QTableWidgetItem("discrepancy"))
                        table_widget.item(i,5).setBackground(QColor('yellow'))
                        discrepancy = discrepancy + 1
                    else:
                        table_widget.setItem(i,5, QTableWidgetItem("reconciliation"))
                        reconciliation = reconciliation + 1
                elif isinstance(df_filtered.iat[i, 1],float):
                    no_ERP = no_ERP + 1
                    table_widget.setItem(i,5, QTableWidgetItem("abnormal"))
                    table_widget.setItem(i,4, QTableWidgetItem('NA'))
                    table_widget.item(i,5).setBackground(QColor('red'))

                elif df_filtered.iat[i, 1] not in SQL_stock_dict:
                    not_in_database = not_in_database + 1
                    table_widget.setItem(i,4, QTableWidgetItem('NA'))
                    table_widget.setItem(i,5, QTableWidgetItem("abnormal"))
                    table_widget.item(i,5).setBackground(QColor('red'))
                
            msg = f"Ecount庫存項目: {table_widget.rowCount()}種, 資料庫庫存項目: {table_widget.rowCount()-not_in_database}種, \n未建ERP Code: {no_ERP}筆, 數量異常: {discrepancy}筆"
            result_label.setText(msg)
            
        table_widget.resizeColumnsToContents()

    def save_excel():
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getSaveFileName(None, "Save File", "Ecount比對表.xlsx", "Excel Files (*.xlsx);", options=options)

        table_data = []

        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            
            rows = table_widget.rowCount()
            columns = table_widget.columnCount()

            # Create a list to hold all the data
            
            # Iterate over all rows and columns to get the content
            for row in range(rows):
                row_data = []
                for column in range(columns):
                    item = table_widget.item(row, column)  # Get the QTableWidgetItem
                    if item is not None:
                        row_data.append(item.text())  # Get the text from the QTableWidgetItem
                    else:
                        row_data.append('')  # If the cell is empty, append an empty string
                table_data.append(row_data)

        df = pd.DataFrame(table_data)
        df.to_excel(file_path, index=False, header=False)

    def hide_show_rows():
        rows = table_widget.rowCount()
        if (hide_button.text() == "顯示異常項目"):
            hide_button.setText("全部顯示")
            for i in range(rows):
                item = table_widget.item(i, 5)
                if item is not None and item.text() == "reconciliation" or item is not None and item.text() == "abnormal":
                    table_widget.setRowHidden(i, True)

        elif (hide_button.text() == "全部顯示"):
            hide_button.setText("顯示異常項目")
            for i in range(rows):
                item = table_widget.item(i, 5)
                if item is not None and item.text() == "reconciliation":
                    table_widget.setRowHidden(i, False)

    load_button.clicked.connect(load_excel)
    save_button.clicked.connect(save_excel)
    hide_button.clicked.connect(hide_show_rows)
    dialog.exec_()

def operation_search_page():
    clear_widgets(widgets)

    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    from main_page import main_page
    button_back.clicked.connect(main_page)

    button1 = create_button("入庫", "#DDDDDD", 0, 0)
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("出庫", "#DDDDDD", 0, 0)
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("退庫", "#DDDDDD", 0, 0)
    widgets["button3"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button_print = create_button("列印/刪除", "#DDDDDD", 0, 0)
    widgets["button_print"].append(button_print)

    #place global widgets on the grid
    grid.addWidget(widgets["button_print"][-1], 1, 4, 1, 1)

    button_operation = create_button("操作查詢", "#DEF100", 0, 0, qcolor=[157,170,0, 60])
    widgets["button_operation"].append(button_operation)

    #place global widgets on the grid
    grid.addWidget(widgets["button_operation"][-1], 1, 5, 1, 1)

    button_print_export = create_button("匯出Excel", "#DDDDDD", 0, 0, qcolor=[157,170,0, 60])
    widgets["button_print_export"].append(button_print_export)

    #place global widgets on the grid
    grid.addWidget(widgets["button_print_export"][-1], 1, 6, 1, 2, alignment=Qt.AlignCenter)

    table_print = create_table()
    widgets["table_print"].append(table_print)

    grid.addWidget(widgets["table_print"][-1], 2, 2, 7, 6)

    # search
    label1 = create_label("編碼", 0, 0, align='right', width=50)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 2, 0, 1, 1)

    #LineEdit widget
    line_bar1 = create_lineedit(0,0,width=170)
    widgets["line_bar1"].append(line_bar1)
    grid.addWidget(widgets["line_bar1"][-1], 2, 1, 1, 2)

    label2 = create_label("入庫日", 0, 0, align='right', width=60)
    widgets["label2"].append(label2)
    # grid.addWidget(widgets["label2"][-1], 3, 0, 1, 1)

    date_choose1 = create_dateedit(0,0,width = 170)
    widgets["date_choose1"].append(date_choose1)
    # grid.addWidget(widgets["date_choose1"][-1], 3, 1, 1, 2)
    # date_choose1.setSpecialValueText(" ")
    # date_choose1.setDate(QDate().fromString("01/01/0001", "dd/MM/yyyy"))

    lineedit_cover = create_lineedit(0,0,width =150)
    widgets['lineedit_cover'].append(lineedit_cover)
    # grid.addWidget(widgets["lineedit_cover"][-1], 3, 1, 1, 2)

    label5 = create_label("地區", 0, 0, align='right', width=50)
    widgets["label5"].append(label5)
    # grid.addWidget(widgets["label5"][-1], 4, 0, 1, 1)


    combo1 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo1"].append(combo1)
    # grid.addWidget(widgets["combo1"][-1], 4, 1, 1, 2)

    label6 = create_label("倉庫", 0, 0, align='right', width=50)
    widgets["label6"].append(label6)
    # grid.addWidget(widgets["label6"][-1], 5, 0, 1, 1)

    combo2 = create_combobox([],0,0,width=170,font_size=20)
    widgets["combo2"].append(combo2)
    # grid.addWidget(widgets["combo2"][-1], 5, 1, 1, 2)
    combo2.setEditable(True)

    label3 = create_label("操作日", 0, 0, align='right', width=60)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 3, 0, 1, 1)

    date_choose2 = create_dateedit(0,0,width = 170)
    widgets["date_choose2"].append(date_choose2)
    grid.addWidget(widgets["date_choose2"][-1], 3, 1, 1, 2)
    # date_choose1.setSpecialValueText(" ")
    # date_choose1.setDate(QDate().fromString("01/01/0001", "dd/MM/yyyy"))

    lineedit_cover2 = create_lineedit(0,0,width =150)
    widgets['lineedit_cover2'].append(lineedit_cover2)
    grid.addWidget(widgets["lineedit_cover2"][-1], 3, 1, 1, 2)

    label4 = create_label("動作", 0, 0, align='right', width=60)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 4, 0, 1, 1)

    combo3 = create_combobox(["入庫", "出庫", "退庫", "刪除", "轉倉"],0,0,width=170,font_size=20)
    widgets["combo3"].append(combo3)
    grid.addWidget(widgets["combo3"][-1], 4, 1, 1, 2)

    button4 = create_button("搜尋", "#DDDDDD", 0, 0)
    widgets["button4"].append(button4)
    # button4.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["button4"][-1], 5, 0, 1, 2)

    button_ecount = create_button("Ecount比對", "#DDDDDD", 0, 0, width=200)
    widgets["button_ecount"].append(button_ecount)

    grid.addWidget(widgets["button_ecount"][-1], 6, 0, 1, 2)

    # button_printer = create_button("修改", "#DDDDDD", 0, 0)
    # widgets["button_printer"].append(button_printer)
    # button4.setVisible(False)

    #place global widgets on the grid
    # grid.addWidget(widgets["button_printer"][-1], 7, 0, 1, 2)
    #place global widgets on the grid
    
    set_today(lineedit_cover)
    set_today(lineedit_cover2)

    combo1.currentIndexChanged.connect(lambda: update_combo2(combo1.currentText(),combo2)) # 20240724 OK
    button1.clicked.connect(button1_click) # 20240724 OK
    button2.clicked.connect(button2_click) # 20240724 OK
    button3.clicked.connect(button3_click) # 20240724 OK
    button_print.clicked.connect(button_print_click) # 20240724 OK
    button4.clicked.connect(lambda: box_operation_inform(table_print, widgets["line_bar1"][-1].text(), widgets["lineedit_cover"][-1].text(), widgets["combo2"][-1].currentText(), widgets["lineedit_cover2"][-1].text(), widgets["combo3"][-1].currentText())) # 20240724 OK
    button_operation.clicked.connect(button_operation_search_click) # 20240724 OK
    # button_printer.clicked.connect(lambda: Operation_Update(table_print, 100))
    date_choose1.dateChanged.connect(lambda: date_change(lineedit_cover, date_choose1)) # 20240724 OK
    date_choose2.dateChanged.connect(lambda: date_change(lineedit_cover2, date_choose2)) # 20240724 OK
    button_print_export.clicked.connect(export_to_excel)
    button_ecount.clicked.connect(Ecount_compare)
