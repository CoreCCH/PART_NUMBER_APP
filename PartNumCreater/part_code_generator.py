import Name_Rule_SQL_handler
from component import main_page_widgets, create_table, create_lineedit, clear_widgets, create_button, create_label, create_combobox, grid, show_alert
from component import part_code_generate_page_widget as widgets
from component import part_code_generate_page_hide_widget as hide_widgets
from PyQt5.QtCore import Qt
import SQL_handler
import execl_handle
from PyQt5.QtWidgets import QTableWidgetItem, QTextEdit, QDialog, QVBoxLayout, QFileDialog
from datetime import datetime

output_data = {}
Pannel = None

def button1_click():
    part_code_generator_page()

def button2_click():
    create_module_page()

def button3_click():
    add_selection_page()

def back_to_main_page():
    from main_page import main_page
    Pannel.hide()
    Pannel.setVisible(False)
    main_page()

def combo_part_item_change(item, category, type, item_combo, category_combo, type_combo, _hide_widgets):
    category_combo.setDisabled(False)

    item_category_list_table = Name_Rule_SQL_handler.fetch_data_from_table({"Item_Name":item}, 'item_correspond_category')

    if isinstance(item_category_list_table, str):
        show_alert(f'讀取命名規則時發生錯誤: {item_category_list_table}')
        category_list= [item[1] for item in item_category_list_table]
    else: 
        category_list = [item[1] for item in item_category_list_table]
        category_combo.clear()
        category_combo.addItems(category_list)
        category_combo.view().setMinimumWidth(300)

    category = category_combo.currentText()

    type_list_table = Name_Rule_SQL_handler.fetch_data_from_table({"Item_Name":item, "Category_Name":category}, table_name= "type_list")

    if isinstance(type_list_table, str):
        show_alert(f"讀取命名規則時發生錯誤: {type_list_table}")
        return
    
    type_list= [item[3] for item in type_list_table]

    type_combo.clear()
    type_combo.addItems(type_list)
    type_combo.setDisabled(False)
    type = type_combo.currentText()

    col_module = Name_Rule_SQL_handler.fetch_data_from_table({"NUM1": item, "NUM2": category if type == "" else category + "_" + type}, table_name= "col_module")
    # print(f"Item Change:{item}, {category}, {type}")
    
    correspond_table = Name_Rule_SQL_handler.get_table_data('correspond_table')
    correspond_dict = {key: value for key, value in correspond_table}


    if isinstance(col_module, str):
        show_alert(f"讀取命名規則時發生錯誤: {col_module}")
        return
    elif col_module == []:
        for num in range(2, 11):
            _hide_widgets["label_NUM"+str(num+1)][-1].setVisible(False)
            _hide_widgets["combo_NUM"+str(num+1)][-1].setVisible(False)
    else:
        for num in range(2, len(col_module[0])-1):
            if col_module[0][num] != None:
                _hide_widgets["label_NUM"+str(num+1)][-1].setText(col_module[0][num])
                _hide_widgets["label_NUM"+str(num+1)][-1].setVisible(True)
                _hide_widgets["combo_NUM"+str(num+1)][-1].setVisible(True)
                
                data = Name_Rule_SQL_handler.get_table_data(correspond_dict[col_module[0][num]])
                

                if col_module[0][num] == "種類":
                    data_list= [val[3] for val in data if val[1] == item and val[2] == category]
                else:
                    if(len(data[0]) == 2):
                        data_list= [val[1] for val in data]
                    else:
                        if type != "":
                            data_list= [val[4] for val in data if val[1] == item and val[2] == category and val[3] == type]
                        else:
                            data_list= [val[4] for val in data if val[1] == item and val[2] == category]
                _hide_widgets["combo_NUM"+str(num+1)][-1].clear()
                _hide_widgets["combo_NUM"+str(num+1)][-1].addItems(data_list)
                _hide_widgets["combo_NUM"+str(num+1)][-1].view().setMinimumWidth(500)
            else:
                _hide_widgets["label_NUM"+str(num+1)][-1].setVisible(False)
                _hide_widgets["combo_NUM"+str(num+1)][-1].setVisible(False)
    
        if col_module[0][2] == "種類":
            _hide_widgets["combo_NUM3"][-1].setDisabled(True)
            _hide_widgets["combo_NUM3"][-1].setCurrentText(type)
        elif col_module[0][2] != "種類":
            _hide_widgets["combo_NUM3"][-1].setDisabled(False)

        if item == "SMT" and category == "電容":
            _hide_widgets["label_NUM12"][-1].setText("特殊規格")
            _hide_widgets["combo_NUM12"][-1].clear()
            data = Name_Rule_SQL_handler.fetch_data_from_table({"Item_Name": "SMT", "Category_Name": "電容"},'percentage_list')
            # print(data)
            unique_data = set(item[3] for item in data)
            unique_data_list = list(unique_data)
            _hide_widgets["combo_NUM12"][-1].addItems(unique_data_list)
            _hide_widgets["label_NUM12"][-1].setVisible(True)
            _hide_widgets["combo_NUM12"][-1].setVisible(True)
        else:
            _hide_widgets["label_NUM12"][-1].setVisible(False)
            _hide_widgets["combo_NUM12"][-1].setVisible(False)

def combo_table_item_change(item, category, type, item_combo, category_combo, type_combo, _hide_widgets):
    category_combo.setDisabled(False)

    item_category_list_table = Name_Rule_SQL_handler.fetch_data_from_table({"Item_Name":item}, 'item_correspond_category')

    if isinstance(item_category_list_table, str):
        show_alert(f'讀取命名規則時發生錯誤: {item_category_list_table}')
        category_list= [item[1] for item in item_category_list_table]
    else: 
        category_list = [item[1] for item in item_category_list_table]
        category_combo.clear()
        category_combo.addItems(category_list)
        category_combo.view().setMinimumWidth(300)

    category = category_combo.currentText()

def combo_part_category_change(item, category, type, item_combo, category_combo, type_combo, _hide_widgets):
    if item == "": return
    if category == "": return
    
    type_list_table = Name_Rule_SQL_handler.fetch_data_from_table({"Item_Name":item, "Category_Name":category}, table_name="type_list")

    if isinstance(type_list_table, str):
        show_alert(f"讀取命名規則時發生錯誤: {type_list_table}")
        return
    
    type_list= [item[3] for item in type_list_table]

    type_combo.clear()
    type_combo.addItems(type_list)
    type_combo.setDisabled(False)
    type = type_combo.currentText()

    col_module = Name_Rule_SQL_handler.fetch_data_from_table({"NUM1": item, "NUM2": category if type == "" else category + "_" + type}, table_name= "col_module")
    correspond_table = Name_Rule_SQL_handler.get_table_data('correspond_table')
    correspond_dict = {key: value for key, value in correspond_table}

    if isinstance(col_module, str):
        show_alert(f"讀取命名規則時發生錯誤: {col_module}")
        return
    elif col_module == []:
        for num in range(2, 11):
            _hide_widgets["label_NUM"+str(num+1)][-1].setVisible(False)
            _hide_widgets["combo_NUM"+str(num+1)][-1].setVisible(False)
    else:
        for num in range(2, len(col_module[0])-1):
            if col_module[0][num] != None:
                _hide_widgets["label_NUM"+str(num+1)][-1].setText(col_module[0][num])
                _hide_widgets["label_NUM"+str(num+1)][-1].setVisible(True)
                _hide_widgets["combo_NUM"+str(num+1)][-1].setVisible(True)

                data = Name_Rule_SQL_handler.get_table_data(correspond_dict[col_module[0][num]])
                if col_module[0][num] == "種類":
                    data_list= [val[3] for val in data if val[1] == item and val[2] == category]
                else:
                    if(len(data[0]) == 2):
                        data_list= [val[1] for val in data]
                    else:
                        if type != "":
                            data_list= [val[4] for val in data if val[1] == item and val[2] == category and val[3] == type]
                        else:
                            data_list= [val[4] for val in data if val[1] == item and val[2] == category]

                _hide_widgets["combo_NUM"+str(num+1)][-1].clear()
                _hide_widgets["combo_NUM"+str(num+1)][-1].addItems(data_list)
                _hide_widgets["combo_NUM"+str(num+1)][-1].view().setMinimumWidth(500)
            else:
                _hide_widgets["label_NUM"+str(num+1)][-1].setVisible(False)
                _hide_widgets["combo_NUM"+str(num+1)][-1].setVisible(False)
    
        if col_module[0][2] == "種類":
            _hide_widgets["combo_NUM3"][-1].setDisabled(True)
            _hide_widgets["combo_NUM3"][-1].setCurrentText(type)
        elif col_module[0][2] != "種類":
            _hide_widgets["combo_NUM3"][-1].setDisabled(False)

        if item == "SMT" and category == "電容":
            _hide_widgets["label_NUM12"][-1].setText("特殊規格")
            _hide_widgets["combo_NUM12"][-1].clear()
            data = Name_Rule_SQL_handler.fetch_data_from_table({"Item_Name": "SMT", "Category_Name": "電容"},'percentage_list')
            unique_data = set(item[3] for item in data)
            unique_data_list = list(unique_data)
            _hide_widgets["combo_NUM12"][-1].addItems(unique_data_list)
            _hide_widgets["label_NUM12"][-1].setVisible(True)
            _hide_widgets["combo_NUM12"][-1].setVisible(True)
        else:
            _hide_widgets["label_NUM12"][-1].setVisible(False)
            _hide_widgets["combo_NUM12"][-1].setVisible(False)

def combo_table_category_change(item, category, type, item_combo, category_combo, type_combo, _hide_widgets):
    if item == "": return
    if category == "": return
    
    type_list_table = Name_Rule_SQL_handler.fetch_data_from_table({"Item_Name":item, "Category_Name":category}, table_name="type_list")

    if isinstance(type_list_table, str):
        show_alert(f"讀取命名規則時發生錯誤: {type_list_table}")
        return
    
    type_list= [item[3] for item in type_list_table]

    type_combo.clear()
    type_combo.addItems(type_list)
    type_combo.setDisabled(False)
    type = type_combo.currentText()

def combo_part_type_change(item, category, type, item_combo, category_combo, type_combo, _hide_widgets):
    if item == "": return
    if category == "": return
    if type == "": return

    col_module = Name_Rule_SQL_handler.fetch_data_from_table({"NUM1": item, "NUM2": category + "_" + type}, table_name= "col_module")
    correspond_table = Name_Rule_SQL_handler.get_table_data('correspond_table')
    correspond_dict = {key: value for key, value in correspond_table}

    if isinstance(col_module, str):
        show_alert(f"讀取命名規則時發生錯誤: {col_module}")
        return
    elif col_module == []:
        for num in range(2, 11):
            _hide_widgets["label_NUM"+str(num+1)][-1].setVisible(False)
            _hide_widgets["combo_NUM"+str(num+1)][-1].setVisible(False)
    else:
        # print(col_module[0])
        for num in range(2, len(col_module[0])-1):
            if col_module[0][num] != None:
                _hide_widgets["label_NUM"+str(num+1)][-1].setText(col_module[0][num])
                _hide_widgets["label_NUM"+str(num+1)][-1].setVisible(True)
                _hide_widgets["combo_NUM"+str(num+1)][-1].setVisible(True)
            
                data = Name_Rule_SQL_handler.get_table_data(correspond_dict[col_module[0][num]])
                if col_module[0][num] == "種類":
                    data_list= [val[3] for val in data if val[1] == item and val[2] == category]
                else:
                    if(len(data[0]) == 2):
                        data_list= [val[1] for val in data]
                    else:
                        data_list= [val[4] for val in data if val[1] == item and val[2] == category and val[3] == type]

                _hide_widgets["combo_NUM"+str(num+1)][-1].clear()
                _hide_widgets["combo_NUM"+str(num+1)][-1].addItems(data_list)
                _hide_widgets["combo_NUM"+str(num+1)][-1].view().setMinimumWidth(500)
            else:
                _hide_widgets["label_NUM"+str(num+1)][-1].setVisible(False)
                _hide_widgets["combo_NUM"+str(num+1)][-1].setVisible(False)
                _hide_widgets["combo_NUM"+str(num+1)][-1].clear()
    
        if col_module[0][2] == "種類":
            _hide_widgets["combo_NUM3"][-1].setDisabled(True)
            _hide_widgets["combo_NUM3"][-1].setCurrentText(type)
        elif col_module[0][2] != "種類":
            _hide_widgets["combo_NUM3"][-1].setDisabled(False)

def create_erp_code(item, category, type, line_bar_pn, lineEdit3,_widgets, _hide_widgets):
    
    ERP_Code = ''
    material_name = ''

    if _widgets["line_bar_pn"][-1].text() == "":
        show_alert("請填入Part Number")
        return
    
    if item == "":
        show_alert("未選擇項目")
        return
    else: 
        ERP_Code += Name_Rule_SQL_handler.fetch_data_from_table({"Name":item},"item_list")[0][0]    

    if category == "":
        show_alert("未選擇種類1")
        return
    else: 
        ERP_Code += Name_Rule_SQL_handler.fetch_data_from_table({"Name":category},"category_list")[0][0]

    col_module = Name_Rule_SQL_handler.fetch_data_from_table({"NUM1": item, "NUM2": category if type == "" else category + "_" + type}, table_name= "col_module")
    correspond_table = Name_Rule_SQL_handler.get_table_data('correspond_table')
    correspond_dict = {key: value for key, value in correspond_table}

    for num in range(2, len(col_module[0])-1):
        if(col_module[0][num] != None):
            data = Name_Rule_SQL_handler.get_table_data(correspond_dict[col_module[0][num]])
            # print(data)
            if(len(data[0]) == 2):
                data_list= [val[0] for val in data if val[1] == _hide_widgets["combo_NUM"+str(num+1)][-1].currentText()]
            
            elif(len(data[0]) == 5):
                data_list= [val[0] for val in data if val[1] == item and val[2] == category and val[3] == _hide_widgets["combo_NUM"+str(num+1)][-1].currentText()]
                
            elif(item == "SMT" and category == "電容" and col_module[0][num] == "%數"):
                
                data_list= [val[0] for val in data if val[1] == item and val[2] == category and val[3] == _hide_widgets["combo_NUM12"][-1].currentText() and val[4] == _hide_widgets["combo_NUM"+str(num+1)][-1].currentText()]
            
            elif(type == ""):
                data_list= [val[0] for val in data if val[1] == item and val[2] == category and val[4] == _hide_widgets["combo_NUM"+str(num+1)][-1].currentText()]
            
            else:
                # print(data)
                # print(f"value:{_hide_widgets["combo_NUM"+str(num+1)][-1].currentText()}")
                data_list= [val[0] for val in data if val[1] == item and val[2] == category and val[3] == type and val[4] == _hide_widgets["combo_NUM"+str(num+1)][-1].currentText()]
                
            ERP_Code = ERP_Code + data_list[0]

    if "零件名稱" == col_module[0][7]:
        material_name = _hide_widgets["combo_NUM8"][-1].currentText().split('(')[0]
    else:
        for num in range(2, len(col_module[0])-4):
            if(col_module[0][num] != None):
                if col_module[0][num] == "%數" and item == "SMT" and category == "電容":
                    material_name = material_name + f"{_hide_widgets["combo_NUM12"][-1].currentText()} {_hide_widgets["combo_NUM"+str(num+1)][-1].currentText()},"
                else:
                    material_name = material_name + _hide_widgets["combo_NUM"+str(num+1)][-1].currentText() + ','

  
    _widgets["line_bar2"][-1].setText(ERP_Code)
    lineEdit3.setText(material_name)

    current_time = datetime.now()

    if(len(ERP_Code) != 11):
        show_alert("生成編碼長度有誤")
        return

    output_df = SQL_handler.get_table_data('Material')
    if isinstance(output_df, str):
        show_alert((f"Error: {output_df}"))
        return

    check_buffer = [output_df[i][0] for i in range(len(output_df))] if output_df is not None else []

    #type = _hide_widgets["combo_NUM3"][-1].currentText() ??
    if type == "": type = None

    if (ERP_Code not in check_buffer):
        if (ERP_Code not in output_data):
            if (_hide_widgets["label_NUM8"][-1].text()== "零件名稱"):
                output_data.update({ERP_Code:["", material_name, item, category, type, _hide_widgets["combo_NUM4"][-1].currentText(), 
                                            f"{_hide_widgets["combo_NUM5"][-1].currentText()}{_hide_widgets["combo_NUM6"][-1].currentText()}{_hide_widgets["combo_NUM7"][-1].currentText()}{_hide_widgets["combo_NUM8"][-1].currentText()}", 
                                            None, _hide_widgets["combo_NUM9"][-1].currentText(), _hide_widgets["combo_NUM11"][-1].currentText(), 
                                            str(current_time.date())+" "+f"{current_time.hour}:{current_time.minute}:{current_time.second}", line_bar_pn.text()]})
            else: 
                output_data.update({ERP_Code:["", material_name, item, category, type, _hide_widgets["combo_NUM4"][-1].currentText(), 
                                            f"{_hide_widgets["combo_NUM5"][-1].currentText()}{_hide_widgets["combo_NUM6"][-1].currentText()}{_hide_widgets["combo_NUM7"][-1].currentText()}", 
                                            _hide_widgets["combo_NUM8"][-1].currentText(), _hide_widgets["combo_NUM9"][-1].currentText(), _hide_widgets["combo_NUM11"][-1].currentText(), 
                                            str(current_time.date())+" "+f"{current_time.hour}:{current_time.minute}:{current_time.second}", line_bar_pn.text()]})
            # part_code += 1
            global Pannel
            Pannel.append(ERP_Code)
            line_bar_pn.clear()

            
    else:
        show_alert("已存在料號: "+ ERP_Code)

def export_data_to_SQL():
    from longin_page import user
    for part_number, values in list(output_data.items()):
        data = [part_number] + values + [user]
        _result = SQL_handler.add_data_to_Material_table(data)
        if (_result == True):
            del output_data[part_number]
    
    global Pannel
    Pannel.clear()
    for part_number in output_data:
        Pannel.append(part_number)

def insert_row_to_table(table):
    new_row_index = table.rowCount()
    table.insertRow(new_row_index)

    for col in range(table.columnCount()):
        item = QTableWidgetItem()
        if col == table.columnCount() - 1:
            # Make the item in the last column not editable
            if col != 1:
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        table.setItem(new_row_index, col, item)

def update_row_to_table(col_name, table, _hide_widget, Item = "", Category = "", Type = ""):
    table_list = Name_Rule_SQL_handler.fetch_data_from_table({"Col_Name": col_name},'correspond_table')
    table_list_name = table_list[0][1]

    columns = Name_Rule_SQL_handler.get_column_names(table_list_name)

    key = columns[-1]

    all_data = []
    for row in range(table.rowCount()):
        row_data = []
        
        for column in range(table.columnCount()):
            item = table.item(row, column)
            if column == 5:
                row_data.append(int(item.text()) if (item != None and item.text() != '') else None)
            else:
                row_data.append(item.text() if (item != None and item.text() != '' and item.text() != 'None') else None)
        all_data.append(row_data)

    all_data = [x for x in all_data if x[0] is not None and x[1] is not None]

    # 使用集合來追踪已見過的組合
    seen = set()
    unique_items = []

    if(len(all_data[0]) >= 5):
        for item in all_data:
            if len(all_data[0]) == 5:
                key = (item[1], item[2], item[3])
            else:
                key = (item[1], item[2], item[3], item[4])
            
            if key not in seen:
                unique_items.append(item)
                seen.add(key)
            else: 
                # print(f"已移除{item}資料，因項目重複")
                show_alert(f"已移除{item}資料，因項目重複")
    else:
        unique_items = all_data

    status = Name_Rule_SQL_handler.insert_or_update_table_data(table_name= table_list_name, columns= columns, data= unique_items, unique_key= key)
 
    if status != True:
        show_alert(f"更新欄位失敗: {status}")
        return
    else:
        show_alert("更新成功!!", "通知")

    show_input(col_name, table, _hide_widget, Item, Category, Type)

def del_selected_data(col_name, table, _hide_widget, Item = "", Category = "", Type = ""):
    table_list = Name_Rule_SQL_handler.fetch_data_from_table({"Col_Name": col_name},'correspond_table')
    table_list_name = table_list[0][1]

    columns = Name_Rule_SQL_handler.get_column_names(table_list_name)

    key = columns[-1]

    selected_indices = table.selectionModel().selectedIndexes()
    
    selected_rows = set(index.row() for index in selected_indices)
    
    key_data = []

    for row in selected_rows:
        for col in range(table.columnCount()):
            item = table.item(row, col)
            if col == (table.columnCount() - 1) and (table.columnCount() - 1) == 5:
                key_data.append(int(item.text()))
            elif col == (table.columnCount() - 1):
                key_data.append(item.text())
        
    status = Name_Rule_SQL_handler.delete_by_keys(table_name= table_list_name, key_column=key, key_values=key_data)

    if status != True:
        show_alert(f"刪除欄位失敗: {status}")
        return
    
    show_input(col_name, table, _hide_widget, Item, Category, Type)
 
def save_to_excel(data):
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    file_path, _ = QFileDialog.getSaveFileName(None, "Save File", "output.xlsx", "Excel Files (*.xlsx);", options=options)
    if file_path:
        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'
        status = execl_handle.export_data_to_excel(data, file_path)

        if status != True:
            show_alert(f"儲存至{file_path}時發生錯誤")
        else:
            show_alert(f"儲存{file_path}成功!!", "通知")

def show_input(col_name, table, _hide_widget, Item = "", Category = "", Type = ""):
    table.clear()
    _hide_widget['combo_NUM3'][-1].clear()
    _hide_widget['combo_NUM4'][-1].clear()
    _hide_widget['combo_NUM5'][-1].clear()

    filter = {}
    if Item != "":
        filter.update({"item_Name":Item})
    if Category != "":
        filter.update({"Category_Name":Category})
    if Type != "":
        filter.update({"Type_Name":Type})

    table_list = Name_Rule_SQL_handler.fetch_data_from_table({"Col_Name": col_name},'correspond_table')
    table_list_name = table_list[0][1]

    columns = Name_Rule_SQL_handler.get_column_names(table_list_name)

    target_table_data = Name_Rule_SQL_handler.get_table_data(table_list_name)
    target_table_data_filter = Name_Rule_SQL_handler.get_table_data(table_list_name, options=filter)

    if (target_table_data_filter == []):
        show_alert("篩選條件有誤")
        show_input(col_name, table, _hide_widget, Item)
        return

    col_len = len(target_table_data_filter[0])

    table.setColumnCount(col_len)


    # header = ["編碼", "名稱"]
    table.setHorizontalHeaderLabels(columns)
    table.setRowCount(len(target_table_data_filter))

    for idx, data in enumerate(target_table_data_filter):
        for num, val in enumerate(data):
            item = QTableWidgetItem(str(val))
            if(num == 5 or (col_len == 2 and num == 1)):
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            table.setItem(idx, num, item)
            
    table.resizeColumnsToContents()

    if col_len == 6:
        item_data = set(item[1] for item in target_table_data)
        item_data_list = list(item_data)

        category_data = set(item[2] for item in target_table_data)
        category_data_list = list(category_data)

        type_data = set(item[3] for item in target_table_data)
        type_data_list = list(type_data)

        _hide_widget['combo_NUM3'][-1].addItems([""]+item_data_list)
        _hide_widget['combo_NUM4'][-1].addItems([""]+category_data_list)
        _hide_widget['combo_NUM5'][-1].addItems([""]+type_data_list)

        _hide_widget["combo_NUM3"][-1].view().setMinimumWidth(280)
        _hide_widget["combo_NUM4"][-1].view().setMinimumWidth(280)
        _hide_widget["combo_NUM5"][-1].view().setMinimumWidth(280)

        _hide_widget['combo_NUM3'][-1].setCurrentText(Item)
        _hide_widget['combo_NUM4'][-1].setCurrentText(Category)
        _hide_widget['combo_NUM5'][-1].setCurrentText(Type)

def renew_table(table):
    table.clear()

    columns = Name_Rule_SQL_handler.get_column_names('col_module')

    table.setHorizontalHeaderLabels(columns)
    table.setColumnCount(len(columns))
    
    col_module_table = Name_Rule_SQL_handler.get_table_data('col_module')
    if isinstance(col_module_table, str):
        show_alert(f'讀取命名規則時發生錯誤: {col_module_table}')
        return

    table.setRowCount(len(col_module_table))

    for idx, data in enumerate(col_module_table):
        for num, val in enumerate(data):
            item = QTableWidgetItem(str(val))
            table.setItem(idx, num, item)
            
    table.resizeColumnsToContents()

    # print(col_module_table)

def add_module(table,_widget, _hide_widget):
    if _widget['combo_item'][-1].currentText() == '':
        show_alert("請選擇項目")
        return
    if _widget['combo_category'][-1].currentText() == '':
        show_alert("請選擇種類1")
        return

    # print(_widget['combo_item'][-1].currentText())
    # print(_widget['combo_category'][-1].currentText())
    # print(_widget['combo_type'][-1].currentText())
    # print(_hide_widget['combo_NUM3'][-1].currentText())

    NUM1 = _widget['combo_item'][-1].currentText()
    NUM2 = _widget['combo_category'][-1].currentText() if _widget['combo_type'][-1].currentText() == "" else f"{_widget['combo_category'][-1].currentText()}_{_widget['combo_type'][-1].currentText()}"
    NUM3 = "種類" if _widget['combo_type'][-1].currentText() != "" else _hide_widget['combo_NUM3'][-1].currentText()
    if NUM3 == "":
        NUM3 = None
    NUM4 = _hide_widget['combo_NUM4'][-1].currentText() if _hide_widget['combo_NUM4'][-1].currentText() != "" else None
    NUM5 = _hide_widget['combo_NUM5'][-1].currentText() if _hide_widget['combo_NUM5'][-1].currentText() != "" else None
    NUM6 = _hide_widget['combo_NUM6'][-1].currentText() if _hide_widget['combo_NUM6'][-1].currentText() != "" else None
    NUM7 = _hide_widget['combo_NUM7'][-1].currentText() if _hide_widget['combo_NUM7'][-1].currentText() != "" else None
    NUM8 = _hide_widget['combo_NUM8'][-1].currentText() if _hide_widget['combo_NUM8'][-1].currentText() != "" else None
    NUM9 = _hide_widget['combo_NUM9'][-1].currentText() if _hide_widget['combo_NUM9'][-1].currentText() != "" else None
    NUM10 = _hide_widget['combo_NUM10'][-1].currentText() if _hide_widget['combo_NUM10'][-1].currentText() != "" else None
    NUM11 = _hide_widget['combo_NUM11'][-1].currentText() if _hide_widget['combo_NUM11'][-1].currentText() != "" else None

    data = [NUM1, NUM2, NUM3, NUM4, NUM5, NUM6, NUM7, NUM8, NUM9, NUM10, NUM11, None]

    status = Name_Rule_SQL_handler.add_data_to_col_module_table(data)

    if status != True:
        show_alert(f"新增模組時發生錯誤: {status}")

    renew_table(table)

def show_material():
    dialog = QDialog()
    dialog.setWindowTitle("顯示結果")
    dialog.resize(1050,550)

    layout = QVBoxLayout()

    datatable = create_table()

    header = ["ERP Code", "ECount", "品項名稱", "項目","總類", "尺寸/種類", "%數","容值/阻值/名稱", "電壓", "製造商", "供應商", "生成時間", "PartNumber"]

    datatable.setColumnCount(len(header))
    datatable.setHorizontalHeaderLabels(header)

    SQL_Material = Name_Rule_SQL_handler.get_table_data(table_name='material',order_by='CreatedAt', order= 'DESC' , database_name= 'pcbmanagement')
    # print(SQL_Material)
    datatable.setRowCount(len(SQL_Material))

    for idx, data in enumerate(SQL_Material):
        datatable.setItem(idx, 0, QTableWidgetItem(SQL_Material[idx][0]))
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
        time_str = SQL_Material[idx][11].strftime("%Y-%m-%d %H:%M:%S")
        datatable.setItem(idx, 11, QTableWidgetItem(time_str))
        datatable.setItem(idx, 12, QTableWidgetItem(SQL_Material[idx][12]))

    datatable.resizeColumnsToContents()

    layout.addWidget(datatable)

    button = create_button("下載資料","#DDDDDD",0,0, width=150)

    button.clicked.connect(lambda: save_to_excel(SQL_Material))

    layout.addWidget(button, alignment= Qt.AlignRight)

    dialog.setLayout(layout)
        
    dialog.exec_()

def part_code_generator_page():
    clear_widgets(widgets)
    clear_widgets(hide_widgets)
    clear_widgets(main_page_widgets)

    item_list_table = Name_Rule_SQL_handler.get_table_data('item_list')
    if isinstance(item_list_table, str):
        show_alert(f'讀取命名規則時發生錯誤: {item_list_table}')
        item_list= []
    else: 
        item_list = [item[1] for item in item_list_table]

    # category_list_table = Name_Rule_SQL_handler.get_table_data('category_list')
    # if isinstance(category_list_table, str):
    #     show_alert(f'讀取命名規則時發生錯誤: {category_list_table}')
    #     category_list= [item[1] for item in category_list_table]
    # else: 
    #     category_list = [item[1] for item in category_list_table]


    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    button_back.clicked.connect(back_to_main_page)

    button1 = create_button("新增編碼頁", "#DEF100", 0, 0)
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("新增模組頁", "#DDDDDD", 0, 0)
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("更新欄位頁", "#DDDDDD", 0, 0)
    widgets["button3"].append(button3)

    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button1.clicked.connect(button1_click) # 20240724 OK
    button2.clicked.connect(button2_click) # 20240724 OK
    button3.clicked.connect(button3_click) # 20240724 OK

    label_item = create_label("項目", 0, 0, align="left")
    widgets["label_item"].append(label_item)
    grid.addWidget(widgets["label_item"][-1], 2, 1, 1, 1, alignment=Qt.AlignLeft)

    combo_item = create_combobox(item_list,0,0)
    widgets["combo_item"].append(combo_item)
    grid.addWidget(widgets["combo_item"][-1], 2, 1, 1, 2, alignment=Qt.AlignHCenter)
    combo_item.view().setMinimumWidth(combo_item.minimumSizeHint().width()+20)
    # widgets["combo_item"][-1].setCurrentIndex(part_choose)
    
    label_category = create_label("種類1", 0, 0, align="left")
    widgets["label_category"].append(label_category)
    grid.addWidget(widgets["label_category"][-1], 3, 1, 1, 1, alignment=Qt.AlignLeft)

    combo_category = create_combobox([],0,0)
    widgets["combo_category"].append(combo_category)
    grid.addWidget(widgets["combo_category"][-1], 3, 1, 1, 2, alignment=Qt.AlignHCenter)
    combo_category.setDisabled(True)
    combo_category.setEditable(True)

    label_type = create_label("種類2", 0, 0, align="left")
    widgets["label_type"].append(label_type)
    grid.addWidget(widgets["label_type"][-1], 4, 1, 1, 1, alignment=Qt.AlignLeft)

    combo_type = create_combobox([],0,0)
    widgets["combo_type"].append(combo_type)
    grid.addWidget(widgets["combo_type"][-1], 4, 1, 1, 2, alignment=Qt.AlignHCenter)
    combo_type.setDisabled(True)

    label_NUM3 = create_label("", 0, 0, align="right", width= 85)
    hide_widgets["label_NUM3"].append(label_NUM3)
    grid.addWidget(hide_widgets["label_NUM3"][-1], 2, 2, 1, 1, alignment=Qt.AlignRight)
    label_NUM3.setVisible(False)

    combo_NUM3 = create_combobox([],0,0)
    hide_widgets["combo_NUM3"].append(combo_NUM3)
    grid.addWidget(hide_widgets["combo_NUM3"][-1], 2, 3, 1, 1, alignment=Qt.AlignLeft)
    combo_NUM3.setVisible(False)

    label_NUM4 = create_label("", 0, 0, align="right", width= 85)
    hide_widgets["label_NUM4"].append(label_NUM4)
    grid.addWidget(hide_widgets["label_NUM4"][-1], 3, 2, 1, 1, alignment=Qt.AlignRight)
    label_NUM4.setVisible(False)

    combo_NUM4 = create_combobox([],0,0)
    hide_widgets["combo_NUM4"].append(combo_NUM4)
    grid.addWidget(hide_widgets["combo_NUM4"][-1], 3, 3, 1, 1, alignment=Qt.AlignLeft)
    combo_NUM4.setVisible(False)

    label_NUM5 = create_label("", 0, 0, align="right", width= 85)
    hide_widgets["label_NUM5"].append(label_NUM5)
    grid.addWidget(hide_widgets["label_NUM5"][-1], 4, 2, 1, 1, alignment=Qt.AlignRight)
    label_NUM5.setVisible(False)

    combo_NUM5 = create_combobox([],0,0)
    hide_widgets["combo_NUM5"].append(combo_NUM5)
    grid.addWidget(hide_widgets["combo_NUM5"][-1], 4, 3, 1, 1, alignment=Qt.AlignLeft)
    combo_NUM5.setVisible(False)

    label_NUM6 = create_label("", 0, 0, align="right", width= 85)
    hide_widgets["label_NUM6"].append(label_NUM6)
    grid.addWidget(hide_widgets["label_NUM6"][-1], 5, 2, 1, 1, alignment=Qt.AlignRight)
    label_NUM6.setVisible(False)

    combo_NUM6 = create_combobox([],0,0)
    hide_widgets["combo_NUM6"].append(combo_NUM6)
    grid.addWidget(hide_widgets["combo_NUM6"][-1], 5, 3, 1, 1, alignment=Qt.AlignLeft)
    combo_NUM6.setVisible(False)

    label_NUM7 = create_label("", 0, 0, align="right", width= 85)
    hide_widgets["label_NUM7"].append(label_NUM7)
    grid.addWidget(hide_widgets["label_NUM7"][-1], 6, 2, 1, 1, alignment=Qt.AlignRight)
    label_NUM7.setVisible(False)

    combo_NUM7 = create_combobox([],0,0)
    hide_widgets["combo_NUM7"].append(combo_NUM7)
    grid.addWidget(hide_widgets["combo_NUM7"][-1], 6, 3, 1, 1, alignment=Qt.AlignLeft)
    combo_NUM7.setVisible(False)

    label_NUM8 = create_label("", 0, 0, align="left", width= 85)
    hide_widgets["label_NUM8"].append(label_NUM8)
    grid.addWidget(hide_widgets["label_NUM8"][-1], 2, 4, 1, 1, alignment=Qt.AlignLeft)
    label_NUM8.setVisible(False)

    combo_NUM8 = create_combobox([],0,0)
    hide_widgets["combo_NUM8"].append(combo_NUM8)
    grid.addWidget(hide_widgets["combo_NUM8"][-1], 2, 4, 1, 2, alignment=Qt.AlignRight)
    combo_NUM8.setVisible(False)

    label_NUM9 = create_label("", 0, 0, align="left", width= 85)
    hide_widgets["label_NUM9"].append(label_NUM9)
    grid.addWidget(hide_widgets["label_NUM9"][-1], 3, 4, 1, 1, alignment=Qt.AlignLeft)
    label_NUM9.setVisible(False)

    combo_NUM9 = create_combobox([],0,0)
    hide_widgets["combo_NUM9"].append(combo_NUM9)
    grid.addWidget(hide_widgets["combo_NUM9"][-1], 3, 4, 1, 2, alignment=Qt.AlignRight)
    combo_NUM9.setVisible(False)

    label_NUM10 = create_label("", 0, 0, align="left", width= 85)
    hide_widgets["label_NUM10"].append(label_NUM10)
    grid.addWidget(hide_widgets["label_NUM10"][-1], 4, 4, 1, 1, alignment=Qt.AlignLeft)
    label_NUM10.setVisible(False)

    combo_NUM10 = create_combobox([],0,0)
    hide_widgets["combo_NUM10"].append(combo_NUM10)
    grid.addWidget(hide_widgets["combo_NUM10"][-1], 4, 4, 1, 2, alignment=Qt.AlignRight)
    combo_NUM10.setVisible(False)
    
    label_NUM11 = create_label("", 0, 0, align="left", width= 85)
    hide_widgets["label_NUM11"].append(label_NUM11)
    grid.addWidget(hide_widgets["label_NUM11"][-1], 5, 4, 1, 1, alignment=Qt.AlignLeft)
    label_NUM11.setVisible(False)

    combo_NUM11 = create_combobox([],0,0)
    hide_widgets["combo_NUM11"].append(combo_NUM11)
    grid.addWidget(hide_widgets["combo_NUM11"][-1], 5, 4, 1, 2, alignment=Qt.AlignRight)
    combo_NUM11.setVisible(False)

    label_NUM12 = create_label("", 0, 0, align="left", width= 85)
    hide_widgets["label_NUM12"].append(label_NUM12)
    grid.addWidget(hide_widgets["label_NUM12"][-1], 6, 4, 1, 1, alignment=Qt.AlignLeft)
    label_NUM12.setVisible(False)

    combo_NUM12 = create_combobox([],0,0)
    hide_widgets["combo_NUM12"].append(combo_NUM12)
    grid.addWidget(hide_widgets["combo_NUM12"][-1], 6, 4, 1, 2, alignment=Qt.AlignRight)
    combo_NUM12.setVisible(False)

    label_pn = create_label("PN", 0, 0, align="left")
    widgets["label_pn"].append(label_pn)
    grid.addWidget(widgets["label_pn"][-1], 6, 1, 1, 1)

    line_bar_pn = create_lineedit( 0, 0, width=150)
    widgets["line_bar_pn"].append(line_bar_pn)
    grid.addWidget(widgets["line_bar_pn"][-1], 6, 1, 1, 2, alignment=Qt.AlignCenter)

    combo_type.currentIndexChanged.connect(lambda: combo_part_type_change(combo_item.currentText(), combo_category.currentText(), combo_type.currentText(), combo_type, combo_category, combo_type, hide_widgets))
    combo_category.currentIndexChanged.connect(lambda: combo_part_category_change(combo_item.currentText(), combo_category.currentText(), combo_type.currentText(), combo_type, combo_category, combo_type, hide_widgets))
    combo_item.currentIndexChanged.connect(lambda: combo_part_item_change(combo_item.currentText(), combo_category.currentText(), combo_type.currentText(), combo_type, combo_category, combo_type, hide_widgets))

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

    global Pannel
    Pannel = QTextEdit()
    Pannel.setReadOnly(True)
    Pannel.setStyleSheet("font-size:20px; font-family: Microsoft JhengHei;font-weight: bold;")
    Pannel.setFixedWidth(200)
    grid.addWidget(Pannel, 3, 6, 4, 2, alignment=Qt.AlignHCenter)

    button2.clicked.connect(lambda: create_erp_code(combo_item.currentText(), combo_category.currentText(), combo_type.currentText(), line_bar_pn, lineEdit3, widgets, hide_widgets))

    label12 = create_label("已產生編號:", 0, 0)
    widgets["label12"].append(label12)
    grid.addWidget(widgets["label12"][-1], 2, 6, 1, 2, alignment=Qt.AlignHCenter|Qt.AlignBottom)

    button3 = create_button("匯出至資料庫", "#1F7145", 0, 0, width=200)
    widgets["button_export"].append(button3)

    #place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 7, 6, 1, 2, alignment=Qt.AlignHCenter)

    button_show = create_button("顯示編碼列表",'#262335', 0, 0, width=200, qcolor=[83,75,76,180], font= 15, font_color='white')
    widgets["button_show"].append(button_show)

    #place global widgets on the grid
    grid.addWidget(widgets["button_show"][-1], 8, 6, 1, 2, alignment=Qt.AlignHCenter)

    button3.clicked.connect(export_data_to_SQL)
    button_show.clicked.connect(show_material)

def create_module_page():
    # for new module create, user have to insert NUM1 ~ NUM11 Data
    clear_widgets(widgets)
    clear_widgets(hide_widgets)
    clear_widgets(main_page_widgets)

    item_list_table = Name_Rule_SQL_handler.get_table_data('item_list')
    if isinstance(item_list_table, str):
        show_alert(f'讀取命名規則時發生錯誤: {item_list_table}')
        item_list= []
    else: 
        item_list = [item[1] for item in item_list_table]

    col_list_table = Name_Rule_SQL_handler.get_table_data('correspond_table')
    if isinstance(item_list_table, str):
        show_alert(f'讀取命名規則時發生錯誤: {col_list_table}')
        col_list= []
    else: 
        col_list = [""]+[item[0] for item in col_list_table]

    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    button_back.clicked.connect(back_to_main_page)

    button1 = create_button("新增編碼頁", "#DDDDDD", 0, 0)
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("新增模組頁", "#DEF100", 0, 0)
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("更新欄位頁", "#DDDDDD", 0, 0)
    widgets["button3"].append(button3)

    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button1.clicked.connect(button1_click) # 20240724 OK
    button2.clicked.connect(button2_click) # 20240724 OK
    button3.clicked.connect(button3_click) # 20240724 OK

    table = create_table()
    widgets['table'].append(table)
    grid.addWidget(widgets['table'][-1], 2, 4,6,4)

    renew_table(table)

    label_item = create_label("項目", 0, 0, align="left", width= 60)
    widgets["label_item"].append(label_item)
    grid.addWidget(widgets["label_item"][-1], 2, 0, 1, 2, alignment=Qt.AlignCenter)
    # label_NUM3.setVisible(False)

    combo_item = create_combobox(item_list,0,0)
    widgets["combo_item"].append(combo_item)
    grid.addWidget(widgets["combo_item"][-1], 2, 1, 1, 2, alignment=Qt.AlignCenter)

    label_category = create_label("種類1", 0, 0, align="left", width= 60)
    widgets["label_category"].append(label_category)
    grid.addWidget(widgets["label_category"][-1], 3, 0, 1, 2, alignment=Qt.AlignCenter)
    # label_NUM3.setVisible(False)

    combo_category = create_combobox([],0,0)
    widgets["combo_category"].append(combo_category)
    grid.addWidget(widgets["combo_category"][-1], 3, 1, 1, 2, alignment=Qt.AlignCenter)

    label_type = create_label("種類2", 0, 0, align="left", width= 60)
    widgets["label_type"].append(label_type)
    grid.addWidget(widgets["label_type"][-1], 4, 0, 1, 2, alignment=Qt.AlignCenter)
    # label_NUM3.setVisible(False)

    combo_type = create_combobox([],0,0)
    widgets["combo_type"].append(combo_type)
    grid.addWidget(widgets["combo_type"][-1], 4, 1, 1, 2, alignment=Qt.AlignCenter)

    label_NUM3 = create_label("編碼3", 0, 0, align="left", width= 60)
    hide_widgets["label_NUM3"].append(label_NUM3)
    grid.addWidget(hide_widgets["label_NUM3"][-1], 5, 0, 1, 2, alignment=Qt.AlignCenter)
    # label_NUM3.setVisible(False)

    combo_NUM3 = create_combobox(col_list,0,0)
    hide_widgets["combo_NUM3"].append(combo_NUM3)
    grid.addWidget(hide_widgets["combo_NUM3"][-1], 5, 1, 1, 2, alignment=Qt.AlignCenter)

    label_NUM4 = create_label("編碼4", 0, 0, align="left", width= 60)
    hide_widgets["label_NUM4"].append(label_NUM4)
    grid.addWidget(hide_widgets["label_NUM4"][-1], 6, 0, 1, 2, alignment=Qt.AlignCenter)
    # label_NUM3.setVisible(False)

    combo_NUM4 = create_combobox(col_list,0,0)
    hide_widgets["combo_NUM4"].append(combo_NUM4)
    grid.addWidget(hide_widgets["combo_NUM4"][-1], 6, 1, 1, 2, alignment=Qt.AlignCenter)

    label_NUM5 = create_label("編碼5", 0, 0, align="left", width= 60)
    hide_widgets["label_NUM5"].append(label_NUM5)
    grid.addWidget(hide_widgets["label_NUM5"][-1], 7, 0, 1, 2, alignment=Qt.AlignCenter)
    # label_NUM3.setVisible(False)

    combo_NUM5 = create_combobox(col_list,0,0)
    hide_widgets["combo_NUM5"].append(combo_NUM5)
    grid.addWidget(hide_widgets["combo_NUM5"][-1], 7, 1, 1, 2, alignment=Qt.AlignCenter)

    label_NUM6 = create_label("編碼6", 0, 0, align="left", width= 60)
    hide_widgets["label_NUM6"].append(label_NUM6)
    grid.addWidget(hide_widgets["label_NUM6"][-1], 8, 0, 1, 2, alignment=Qt.AlignCenter)
    # label_NUM3.setVisible(False)

    combo_NUM6 = create_combobox(col_list,0,0)
    hide_widgets["combo_NUM6"].append(combo_NUM6)
    grid.addWidget(hide_widgets["combo_NUM6"][-1], 8, 1, 1, 2, alignment=Qt.AlignCenter)

    label_NUM7 = create_label("編碼7", 0, 0, align="left", width= 60)
    hide_widgets["label_NUM7"].append(label_NUM7)
    grid.addWidget(hide_widgets["label_NUM7"][-1], 2, 2, 1, 1, alignment=Qt.AlignRight)
    # label_NUM3.setVisible(False)

    combo_NUM7 = create_combobox(col_list,0,0)
    hide_widgets["combo_NUM7"].append(combo_NUM7)
    grid.addWidget(hide_widgets["combo_NUM7"][-1], 2, 2, 1, 2, alignment=Qt.AlignRight)

    label_NUM8 = create_label("編碼8", 0, 0, align="left", width= 60)
    hide_widgets["label_NUM8"].append(label_NUM8)
    grid.addWidget(hide_widgets["label_NUM8"][-1], 3, 2, 1, 1, alignment=Qt.AlignRight)
    # label_NUM3.setVisible(False)

    combo_NUM8 = create_combobox(col_list,0,0)
    hide_widgets["combo_NUM8"].append(combo_NUM8)
    grid.addWidget(hide_widgets["combo_NUM8"][-1], 3, 2, 1, 2, alignment=Qt.AlignRight)

    label_NUM9 = create_label("編碼9", 0, 0, align="left", width= 60)
    hide_widgets["label_NUM9"].append(label_NUM9)
    grid.addWidget(hide_widgets["label_NUM9"][-1], 4, 2, 1, 1, alignment=Qt.AlignRight)
    # label_NUM3.setVisible(False)

    combo_NUM9 = create_combobox(["廠商"],0,0)
    hide_widgets["combo_NUM9"].append(combo_NUM9)
    grid.addWidget(hide_widgets["combo_NUM9"][-1], 4, 2, 1, 2, alignment=Qt.AlignRight)
    combo_NUM9.setCurrentText("廠商")

    label_NUM10 = create_label("編碼10", 0, 0, align="left", width= 70)
    hide_widgets["label_NUM10"].append(label_NUM10)
    grid.addWidget(hide_widgets["label_NUM10"][-1], 5, 2, 1, 1, alignment=Qt.AlignRight)
    # label_NUM3.setVisible(False)

    combo_NUM10 = create_combobox([],0,0)
    hide_widgets["combo_NUM10"].append(combo_NUM10)
    grid.addWidget(hide_widgets["combo_NUM10"][-1], 5, 2, 1, 2, alignment=Qt.AlignRight)
    combo_NUM10.setEnabled(False)

    label_NUM11 = create_label("編碼11", 0, 0, align="left", width= 70)
    hide_widgets["label_NUM11"].append(label_NUM11)
    grid.addWidget(hide_widgets["label_NUM11"][-1], 6, 2, 1, 1, alignment=Qt.AlignRight)
    # label_NUM3.setVisible(False)

    combo_NUM11 = create_combobox(["供應商"],0,0)
    hide_widgets["combo_NUM11"].append(combo_NUM11)
    grid.addWidget(hide_widgets["combo_NUM11"][-1], 6, 2, 1, 2, alignment=Qt.AlignRight)
    combo_NUM11.setCurrentText("供應商")


    combo_category.currentIndexChanged.connect(lambda: combo_table_category_change(combo_item.currentText(), combo_category.currentText(), combo_type.currentText(), combo_type, combo_category, combo_type, hide_widgets))
    combo_item.currentIndexChanged.connect(lambda: combo_table_item_change(combo_item.currentText(), combo_category.currentText(), combo_type.currentText(), combo_type, combo_category, combo_type, hide_widgets))

    button_export = create_button("新增模組", "#DDDDDD", 0, 0)
    widgets['button_export'].append(button_export)
    grid.addWidget(widgets['button_export'][-1], 8, 2, 1, 2, alignment=Qt.AlignRight)

    button_export.clicked.connect(lambda: add_module(table, widgets, hide_widgets))

def add_selection_page():
    # add selection
    clear_widgets(widgets)
    clear_widgets(hide_widgets)
    clear_widgets(main_page_widgets)

    item_category_list_table = Name_Rule_SQL_handler.get_table_data('correspond_table')
    if isinstance(item_category_list_table, str):
        show_alert(f'讀取命名規則時發生錯誤: {item_category_list_table}')
        item_list= []
    else: 
        item_list = [item[0] for item in item_category_list_table]

    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    button_back.clicked.connect(back_to_main_page)

    button1 = create_button("新增編碼頁", "#DDDDDD", 0, 0)
    widgets["button1"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button1"][-1], 1, 1, 1, 1)

    button2 = create_button("新增模組頁", "#DDDDDD", 0, 0)
    widgets["button2"].append(button2)

    #place global widgets on the grid
    grid.addWidget(widgets["button2"][-1], 1, 2, 1, 1)

    button3 = create_button("更新欄位頁", "#DEF100", 0, 0)
    widgets["button3"].append(button3)

    grid.addWidget(widgets["button3"][-1], 1, 3, 1, 1)

    button1.clicked.connect(button1_click) # 20240724 OK
    button2.clicked.connect(button2_click) # 20240724 OK
    button3.clicked.connect(button3_click) # 20240724 OK

    label_choose = create_label("請選擇要更新的選項:",0,0,align="left",width=300)
    widgets["label_choose"].append(label_choose)
    grid.addWidget(widgets["label_choose"][-1], 2, 1, 1, 2, alignment=Qt.AlignRight)

    combo_choose = create_combobox(item_list,0,0,width=200)
    widgets["combo_choose"].append(combo_choose)
    grid.addWidget(widgets["combo_choose"][-1], 2, 2, 1, 2, alignment=Qt.AlignCenter)

    label_NUM3 = create_label("項目", 0, 0, align="right", width= 85)
    hide_widgets["label_NUM3"].append(label_NUM3)
    grid.addWidget(hide_widgets["label_NUM3"][-1], 3, 1, 1, 1, alignment=Qt.AlignRight)
    # label_NUM3.setVisible(False)

    combo_NUM3 = create_combobox([],0,0)
    hide_widgets["combo_NUM3"].append(combo_NUM3)
    grid.addWidget(hide_widgets["combo_NUM3"][-1], 3, 2, 1, 1, alignment=Qt.AlignLeft)
    # combo_NUM3.setVisible(False)

    label_NUM4 = create_label("種類1", 0, 0, align="right", width= 85)
    hide_widgets["label_NUM4"].append(label_NUM4)
    grid.addWidget(hide_widgets["label_NUM4"][-1], 4, 1, 1, 1, alignment=Qt.AlignRight)
    # label_NUM3.setVisible(False)

    combo_NUM4 = create_combobox([],0,0)
    hide_widgets["combo_NUM4"].append(combo_NUM4)
    grid.addWidget(hide_widgets["combo_NUM4"][-1], 4, 2, 1, 1, alignment=Qt.AlignLeft)
    # combo_NUM3.setVisible(False)

    label_NUM5 = create_label("種類2", 0, 0, align="right", width= 85)
    hide_widgets["label_NUM5"].append(label_NUM5)
    grid.addWidget(hide_widgets["label_NUM5"][-1], 5, 1, 1, 1, alignment=Qt.AlignRight)
    # label_NUM3.setVisible(False)

    combo_NUM5 = create_combobox([],0,0)
    hide_widgets["combo_NUM5"].append(combo_NUM5)
    grid.addWidget(hide_widgets["combo_NUM5"][-1], 5, 2, 1, 1, alignment=Qt.AlignLeft)
    # combo_NUM3.setVisible(False)

    table = create_table()
    widgets['table'].append(table)
    grid.addWidget(widgets['table'][-1], 2, 4,6,4)

    combo_choose.currentIndexChanged.connect(lambda: show_input(combo_choose.currentText(), table, hide_widgets))

    button_filter = create_button("篩選", "#FFFDD4", 0, 0)
    widgets['button_output'].append(button_filter)
    grid.addWidget(widgets['button_output'][-1], 6,2,1,1, alignment=Qt.AlignRight)

    button_filter.clicked.connect(lambda: show_input(combo_choose.currentText(), table, hide_widgets, combo_NUM3.currentText(), combo_NUM4.currentText(), combo_NUM5.currentText()))

    button_insert = create_button("新增", "#DDDDDD", 0, 0)
    widgets['button_insert'].append(button_insert)
    grid.addWidget(widgets['button_insert'][-1], 8, 4,1,1, alignment=Qt.AlignCenter)

    button_insert.clicked.connect(lambda: insert_row_to_table(table))

    button_export = create_button("更新", "#DDDDDD", 0, 0)
    widgets['button_export'].append(button_export)
    grid.addWidget(widgets['button_export'][-1], 8, 5,1,2, alignment=Qt.AlignLeft)

    button_export.clicked.connect(lambda: update_row_to_table(combo_choose.currentText(), table, hide_widgets, combo_NUM3.currentText(), combo_NUM4.currentText(), combo_NUM5.currentText()))

    button_delete = create_button("刪除", "#DDDDDD", 0, 0)
    widgets['button_delete'].append(button_delete)
    grid.addWidget(widgets['button_delete'][-1], 8, 6,1,2, alignment=Qt.AlignCenter)

    button_delete.clicked.connect(lambda: del_selected_data(combo_choose.currentText(), table, hide_widgets, combo_NUM3.currentText(), combo_NUM4.currentText(), combo_NUM5.currentText()))