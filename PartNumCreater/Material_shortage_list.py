from component import grid, create_button, create_table, create_lineedit, create_label, show_alert, clear_widgets, create_combobox, main_page_widgets
from component import material_shortage_widgets as widgets
from PyQt5.QtWidgets import QComboBox, QFileDialog,QTabWidget, QTableWidgetItem, QPushButton, QDialog, QTreeWidgetItem, QVBoxLayout, QLabel, QHBoxLayout, QGridLayout,QTreeWidget, QAbstractItemView, QHeaderView, QMessageBox, QWidget, QTableWidget, QSpinBox
from PyQt5.QtCore import Qt, QRegExp
from PyQt5.QtGui import QColor, QFont, QRegExpValidator
from datetime import datetime, timedelta
import execl_handle, Name_Rule_SQL_handler
import SQL_handler
import pandas as pd

filePath = ""

collected_data = {}

def create_line_edits(Vlayout_2, num, _filePath):
    bom_folder_list = SQL_handler.get_table_data(database_name="bommanagement", table_name="folder_list")
    if(isinstance(bom_folder_list, str)):
        show_alert(f"讀取資料夾列表時發生錯誤: {bom_folder_list}")
        bom_folder_list = []

    bom_folder_name_list = [item[0] for item in bom_folder_list]

    # Clear existing QLineEdits
    for i in reversed(range(Vlayout_2.count())): 
        for j in reversed(range(Vlayout_2.itemAt(i).count())):
            Vlayout_2.itemAt(i).itemAt(j).widget().deleteLater()
        Vlayout_2.itemAt(i).deleteLater()

    # Create new QLineEdits
    if (_filePath != []):
        for idx in range(num):
            Hlayout = QHBoxLayout()
            label = create_label(f'請輸入檔案{str(idx+1)}名稱:',0,0,align='right',width=200)
            Hlayout.addWidget(label)
            
            line_edit = create_lineedit(0,0, width=500, font_size=15)
            line_edit.setPlaceholderText(_filePath[idx].split('/')[-1])
            # regex = QRegExp("[a-z0-9]+")
            # validator = QRegExpValidator(regex, line_edit)
            # line_edit.setValidator(validator)
            Hlayout.addWidget(line_edit)

            label = create_label(f"選擇資料夾",0,0,align="right")
            Hlayout.addWidget(label)

            combo_folder = create_combobox(bom_folder_name_list,0,0)
            combo_folder.setEditable(True)
            Hlayout.addWidget(combo_folder)

            Vlayout_2.addLayout(Hlayout)

def get_file_path(Vlayout_2, input1):
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    _filePath, _ = QFileDialog.getOpenFileNames(None,"Open File", "", "Text Files (*.xlsx)", options=options)
    filePath_str =  ','.join(f'"{item}"' for item in _filePath)
    input1.setText(filePath_str)

    create_line_edits(Vlayout_2, len(input1.text().split(',')), _filePath)

def import_excel_file_to_SQL(line_edit, Vlayout_2, dialog):
    from main_page import user
    movtivation_list = []
    path_list = line_edit.text().replace('"','').split(',')
    
    for i in range(Vlayout_2.count()):
        if (Vlayout_2.itemAt(i).itemAt(1).widget().text() != "" and Vlayout_2.itemAt(i).itemAt(1).widget().text() != None and Vlayout_2.itemAt(i).itemAt(3).widget().currentText() != "" and Vlayout_2.itemAt(i).itemAt(3).widget().currentText() != None):
            bom_folder_list = SQL_handler.get_table_data(database_name="bommanagement", table_name="folder_list")
            if(isinstance(bom_folder_list, str)):
                show_alert(f"讀取資料夾列表時發生錯誤: {bom_folder_list}")
                bom_folder_list = []

            SQL_Table_Name = Vlayout_2.itemAt(i).itemAt(1).widget().text()

            SQL_Folder_Name = Vlayout_2.itemAt(i).itemAt(3).widget().currentText().lower()

            df = execl_handle.excel_file_read(path_list[i], execl_handle.read_excel_sheets(path_list[i])[0])
           
            data = execl_handle.read_BOM(df)
            filtered_data = [item for item in data if all(len(str(key)) == 11 for key in item.keys())]

            status = SQL_handler.create_and_populate_table(SQL_Table_Name, filtered_data, database_name='bom')

            if status == True:
                show_alert(f"{SQL_Table_Name} 匯入資料庫成功!!","通知")
                movtivation_list.append(("ADD", "BOM", SQL_Folder_Name, datetime.now(), user))
                dialog.reject()
            else:
                show_alert(f"{SQL_Table_Name} 匯入資料庫失敗: {status}")
                dialog.reject()
                return

            bom_folder_name_list = [item[0] for item in bom_folder_list]
            
            if(SQL_Folder_Name not in bom_folder_name_list):
                
                status = SQL_handler.add_data_to_folder_table(foler_name=SQL_Folder_Name, database_name="bommanagement")
                if status != True:
                    show_alert(f"{SQL_Folder_Name} 匯入資料庫失敗: {status}")
                    dialog.reject()
                    return

                movtivation_list.append(("ADD", "FOLDER", SQL_Folder_Name, datetime.now(), user))
            
            

            status = SQL_handler.add_data_to_folder_BOM_table(Folder_BOM_list=[SQL_Folder_Name, SQL_Table_Name], database_name="bommanagement")
            if status != True:
                show_alert(f"{[SQL_Folder_Name, SQL_Table_Name]} 匯入資料庫失敗: {status}")
                dialog.reject()
                return
            
            
            
    SQL_handler.add_data_to_Bom_Operation_table(movtivation_list, database_name="bommanagement")       
       
def import_excel():
    dialog = QDialog()
    dialog.setWindowTitle("匯入資料")
    dialog.resize(1050,550)

    Vlayout = QVBoxLayout()
    Vlayout_2 = QVBoxLayout()
    Hlayout_1 = QHBoxLayout()

    label1 = create_label("檔案:",0,0, align= "right",width=100, font_size=18)
    input1 = create_lineedit(0,0,width=500, font_size=15)
    button_get_path = create_button("讀取檔案", "#DDDDDD", 0, 0, height= 44,width=100, font=20)
    button_get_path.clicked.connect(lambda: get_file_path(Vlayout_2, input1))
    Hlayout_1.addWidget(label1)
    Hlayout_1.addWidget(input1)
    Hlayout_1.addWidget(button_get_path)
    Vlayout.addLayout(Hlayout_1)
    Vlayout.addLayout(Vlayout_2)

    button_import = create_button("匯入BOM至資料庫","#DDDDDD",0,0,width=300)
    button_import.clicked.connect(lambda: import_excel_file_to_SQL(input1, Vlayout_2, dialog))

    Vlayout.addWidget(button_import, alignment=Qt.AlignRight)

    dialog.setLayout(Vlayout)
        
    dialog.exec_()

def back_to_main_page():
    from main_page import main_page
    main_page()

def get_future_months(start_date, months):
    future_months = []
    for i in range(months):
        future_month = start_date + timedelta(days=(i * 30))
        future_months.append(future_month.strftime("%Y/%m"))
    return future_months

def caculate_all_remander(data):
    accumulated_data = {}

    for item in data:
        for key, value in item.items():
            if key in accumulated_data:
                accumulated_data[key] += value
            else:
                accumulated_data[key] = value

    result = [{key: value} for key, value in accumulated_data.items()]

    return result

def move_item_right(left_tree, right_tree):
    items = left_tree.selectedItems()
    for item in items:
        left_tree.invisibleRootItem().removeChild(item)
        right_tree.invisibleRootItem().addChild(item)

def move_item_left(left_tree, right_tree):
    items = right_tree.selectedItems()
    for item in items:
        right_tree.invisibleRootItem().removeChild(item)
        left_tree.invisibleRootItem().addChild(item)

def hideColumnOnClick(table, column, button):
    # 隱藏被點擊的列
    item_len = column - 19

    if (button.text() == "隱藏\n欄位"):
        for idx in range(1,1+item_len):
            table.setColumnHidden(idx, True)
            table.viewport().update()
        button.setText("顯示\n欄位")
    elif (button.text() == "顯示\n欄位"):
        for idx in range(1,1+item_len):
            table.setColumnHidden(idx, False)
            table.viewport().update()
        button.setText("隱藏\n欄位")

def confirm_selection(linedit_filename, right_tree, dialog, table, table2):
    previous_name = linedit_filename.text()
    selected_items = []
    root = right_tree.invisibleRootItem()
    
    # Traverse through each top-level item
    for i in range(root.childCount()):
        top_level_item = root.child(i)
        
        # Traverse through the child items of each top-level item
        for j in range(top_level_item.childCount()):
            selected_items.append(top_level_item.child(j).text(0))
    
    if not selected_items:
        linedit_filename.setText('請點擊選擇BOM表')
    else:
        linedit_filename.setText(','.join(f'{item}' for item in selected_items))

    if (linedit_filename.text() != previous_name):
        # clear record if selected BOM change
        collected_data.clear()

    dialog.reject()

    table_show(linedit_filename, table, table2)

def delete_selection(linedit_filename, right_tree, left_tree, dialog, table, table2):  
    selected_items = []
    selected_root = []
    right_root = right_tree.selectedItems()
    left_root = left_tree.selectedItems()
    

    for item in right_root:
        if item.parent() is None:
            selected_root.append(item.text(0))
            for i in range(item.childCount()):
                selected_items.append(item.child(i).text(0))
        else:
      
            selected_items.append(item.text(0))
    
    for item in left_root:
        if item.parent() is None:
            selected_root.append(item.text(0))
            for i in range(item.childCount()):
                selected_items.append(item.child(i).text(0))
        else:
        
            selected_items.append(item.text(0))



    message = f"是否刪除 {selected_root} 資料夾以及資料 {set(selected_items)}?"
    reply = QMessageBox.question(dialog, '確認刪除', message, 
                                 QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

    if reply == QMessageBox.Yes:
        from main_page import user
        print(user)

        movtivation_list = []

        if (selected_items != []):
            status = Name_Rule_SQL_handler.drop_table(table_names=list(set(selected_items)), database_name= "bom")
            if status != True:
                show_alert(f"刪除BOM表時發生錯誤: {status}")
                return
            
            for BOM_name in list(set(selected_items)):
                movtivation_list.append(("DELETE", "BOM", BOM_name, datetime.now(), user))
                

            status = Name_Rule_SQL_handler.delete_by_keys(table_name= "folder_correspond_bom", key_column= "BOM_Name", key_values= list(set(selected_items)), database_name= "bommanagement")
            if status != True:
                show_alert(f"刪除BOM表時發生錯誤: {status}")
                return
        
        if (selected_root != []):
            status = Name_Rule_SQL_handler.delete_by_keys(table_name= "folder_list", key_column= "Folder_Name", key_values= selected_root, database_name= "bommanagement")
            if status != True:
                show_alert(f"刪除資料夾時發生錯誤: {status}")
                return
            
            for root_name in selected_root:
                movtivation_list.append(("DELETE", "Folder", root_name, datetime.now(), user))
        
        SQL_handler.add_data_to_Bom_Operation_table(movtivation_list, database_name="bommanagement")
        dialog.reject()

    else:
        # 用戶取消刪除，返回
        return

def showDialog(linedit_filename, table, table2):
    folder_list = SQL_handler.get_table_data(table_name="folder_list", database_name="bommanagement")
    folder_BOM_list = SQL_handler.get_table_data(table_name="folder_correspond_bom", database_name="bommanagement")

    BOM_list = []
    if linedit_filename.text() != "請點擊選擇BOM表":
        BOM_list = linedit_filename.text().split(',')

    dialog = QDialog()
    dialog.setWindowTitle("BOM表輸入")
    dialog.resize(650, 350)

    table_list = SQL_handler.list_tables(database_name="bom")

    if isinstance(table_list, str):
        show_alert(f"讀取table列表時發生異常: {table_list}")
        return
    elif table_list == []:
        show_alert(f"資料庫尚無BOM表資料，請點擊右方[匯入]按鈕將EBOM匯入資料庫")
        import_excel()
        BOM_list = linedit_filename.text().split(',')
        return

    main_layout = QVBoxLayout()
    grid_layout = QGridLayout()

    # 左 TreeWidget
    left_tree = QTreeWidget()
    left_tree.setHeaderLabels(["BOM表清單:"])
    left_tree.setSelectionMode(QAbstractItemView.MultiSelection)

    # 右 TreeWidget
    right_tree = QTreeWidget()
    right_tree.setHeaderLabels(["已選擇BOM表:"])
    right_tree.setSelectionMode(QAbstractItemView.MultiSelection)

    # 建立 left_tree 結構
    for folder in folder_list:
        folder_name = folder[0]
        left_parent_item = QTreeWidgetItem(left_tree, [folder_name])
        left_parent_item.setExpanded(True)

        # 添加子项
        for bom in folder_BOM_list:
            if bom[0] == folder_name and bom[1] not in BOM_list:
                QTreeWidgetItem(left_parent_item, [bom[1]])

    # 建立 right_tree 結構
    right_tree_folders = {folder[0]: QTreeWidgetItem(right_tree, [folder[0]]) for folder in folder_list}
    for right_folder_item in right_tree_folders.values():
        right_folder_item.setExpanded(True)

    for table_name in BOM_list:
        QTreeWidgetItem(right_tree_folders.get(folder_list[0][0]), [table_name])

    def move_item_right(left_tree, right_tree):
        selected_items = left_tree.selectedItems()
        if not selected_items:
            return
        
        for item in selected_items:
            if item.childCount() > 0:
                # Handle moving folders and their children
                folder_name = item.text(0)
                right_folder_item = right_tree.findItems(folder_name, Qt.MatchExactly)
                if right_folder_item:
                    right_folder_item = right_folder_item[0]
                else:
                    # Create folder if it does not exist
                    right_folder_item = QTreeWidgetItem(right_tree, [folder_name])
                    right_folder_item.setExpanded(True)
                
                for i in reversed(range(item.childCount())):
                    child = item.child(i)
                    right_folder_item.addChild(child.clone())
                    item.removeChild(child)
                
                # Remove the folder if it has no more children
                if item.childCount() == 0 and item.parent():
                    item.parent().removeChild(item)
                elif item.childCount() == 0:
                    left_tree.invisibleRootItem().removeChild(item)
            else:
                # Handle moving individual items
                folder_name = item.parent().text(0) if item.parent() else None
                if folder_name:
                    right_folder_item = right_tree.findItems(folder_name, Qt.MatchExactly)
                    if right_folder_item:
                        right_folder_item = right_folder_item[0]
                    else:
                        right_folder_item = QTreeWidgetItem(right_tree, [folder_name])
                        right_folder_item.setExpanded(True)
                
                    right_folder_item.addChild(item.clone())
                    if item.parent():
                        item.parent().removeChild(item)
                    else:
                        left_tree.invisibleRootItem().removeChild(item)

        # Clear the selection
        left_tree.clearSelection()

    def move_item_left(left_tree, right_tree):
        selected_items = right_tree.selectedItems()
        if not selected_items:
            return

        for item in selected_items:
            if item.childCount() > 0:
                # Handle moving folders and their children
                folder_name = item.text(0)
                left_folder_item = left_tree.findItems(folder_name, Qt.MatchExactly)
                if left_folder_item:
                    left_folder_item = left_folder_item[0]
                else:
                    # Create folder if it does not exist
                    left_folder_item = QTreeWidgetItem(left_tree, [folder_name])
                    left_folder_item.setExpanded(True)

                for i in reversed(range(item.childCount())):
                    child = item.child(i)
                    left_folder_item.addChild(child.clone())
                    item.removeChild(child)

                # Remove the folder if it has no more children
                if item.childCount() == 0 and item.parent():
                    item.parent().removeChild(item)
                elif item.childCount() == 0:
                    right_tree.invisibleRootItem().removeChild(item)
            else:
                # Handle moving individual items
                folder_name = item.parent().text(0) if item.parent() else None
                if folder_name:
                    left_folder_item = left_tree.findItems(folder_name, Qt.MatchExactly)
                    if left_folder_item:
                        left_folder_item = left_folder_item[0]
                    else:
                        left_folder_item = QTreeWidgetItem(left_tree, [folder_name])
                        left_folder_item.setExpanded(True)
                
                    left_folder_item.addChild(item.clone())
                    if item.parent():
                        item.parent().removeChild(item)
                    else:
                        right_tree.invisibleRootItem().removeChild(item)

        # Clear the selection
        right_tree.clearSelection()

    # 左箭
    left_button = create_button("<", "#DDDDDD", 0, 0)
    left_button.clicked.connect(lambda: move_item_left(left_tree, right_tree))

    # 右箭
    right_button = create_button(">", "#DDDDDD", 0, 0)
    right_button.clicked.connect(lambda: move_item_right(left_tree, right_tree))

    grid_layout.addWidget(left_tree, 0, 0, 4, 1)
    grid_layout.addWidget(right_button, 2, 1, 1, 1)
    grid_layout.addWidget(left_button, 3, 1, 1, 1)
    grid_layout.addWidget(right_tree, 0, 2, 4, 1)

    button_layout = QHBoxLayout()
    confirm_button = create_button("確認", "#DDDDDD", 0, 0)
    confirm_button.clicked.connect(lambda: confirm_selection(linedit_filename, right_tree, dialog, table, table2))
    button_layout.addWidget(confirm_button)

    delete_button = create_button("刪除", "#DDDDDD", 0, 0)
    delete_button.clicked.connect(lambda: delete_selection(linedit_filename, right_tree, left_tree, dialog, table, table2))
    button_layout.addWidget(delete_button)

    button_layout.addStretch()

    main_layout.addLayout(grid_layout)
    main_layout.addLayout(button_layout)

    dialog.setLayout(main_layout)

    dialog.exec_()

def table_show(linedit_filename, table, table2, button_bom = None):
    table.clear()
    table2.clear()

    if isinstance(linedit_filename, QPushButton):
        if(linedit_filename.text() == "請點擊選擇BOM表"):
            show_alert("請先選擇BOM表")
            return

        table_list = linedit_filename.text().split(',')    
    elif isinstance(linedit_filename, QComboBox):
        data = Name_Rule_SQL_handler.get_table_data(table_name="forecast_list", database_name="forecast", options={"Group_Name":linedit_filename.currentText()})
        table_list = list(set([item[2] for item in data]))
        collected_data.clear()
        for val in data:
            date_str = val[4].strftime('%Y-%m-%d')
            collected_data.update({(val[2],date_str):val[3]})

    # table.setColumnCount(2+len(table_list)+18)
    table.setColumnCount(2+len(table_list))

    header = ["ERP Code"]
    header2 = ['專案名稱']

    text = ""

    for i in range(len(table_list)):
        text = text + (str(i+1) + "." +table_list[i].split('/')[-1].split('_')[0] + " ")
        header.append(table_list[i].split('/')[-1].split('_')[0]+"需量/1pc")
        # header2.append(table_list[i].split('/')[-1].split('_')[0]+"估量")

    header.append("庫存")
    
    future_months = get_future_months(datetime.now(), 18)
    for i in range(len(future_months)):
        # header.append(future_months[i]+"餘料")
        header2.append(future_months[i]+"估量")
    
    
    table.setHorizontalHeaderLabels(header)

    erp_codes_list = []
    quantities_list = []
    final_list = []
    data = []

    for i in range(len(table_list)):
        df = SQL_handler.get_table_data(table_name=table_list[i], database_name='bom')
        print(df)
        for val in df:
            data.append({val[0]:val[1]})
        erp_codes_list.append([list(item.keys())[0] for item in data])
        quantities_list.append([list(item.values())[0] for item in data])
        final_list += [list(item.keys())[0] for item in data]
    
    none_repeat = []
    for item in final_list:
        if item not in none_repeat:
            none_repeat.append(item)

    # table.setRowCount(len(list(set(final_list))))
    table.setRowCount(len(none_repeat))

    combo_list = [none_repeat]

    for i in range(len(table_list)):
        combo_list.append([0] * len(none_repeat))

    for i in range(len(erp_codes_list)):
        for j, item in enumerate(none_repeat):
            if item in erp_codes_list[i]:
                index = erp_codes_list[i].index(item)
                combo_list[i + 1][j] = quantities_list[i][index]

    SQL_stock = SQL_handler.get_table_data(table_name='Inventory')
    
 
    # 初始化空字典
    SQL_stock_dict = {}
  

    # 遍歷資料
    for item in SQL_stock:
        key = item[1]
        quantity = item[2]
        
        if key in SQL_stock_dict:
            SQL_stock_dict[key] += quantity
        else:
            SQL_stock_dict[key] = quantity

    print(SQL_stock_dict)

    for row in range(len(combo_list[0])):
        for col in range(len(combo_list)):
            table.setItem(row, col, QTableWidgetItem(str(combo_list[col][row])))  
            
        count = SQL_stock_dict.get(combo_list[0][row], 0)
        table.setItem(row, col+1, QTableWidgetItem(str(count)))
    
    table2.setColumnCount(19)
    table2.setRowCount(len(table_list))

    for i in range(len(table_list)):
        table2.setItem(i, 0, QTableWidgetItem(table_list[i].split('/')[-1].split('_')[0]))


    table2.setHorizontalHeaderLabels(header2)

    set_header = table2.horizontalHeader()
    set_header.setSectionsClickable(True)

    def sum_quantities_by_month():
        monthly_sums = {}
        # print(f"HI: {collected_data}")
        # Iterate over the items in collected_data
        for (name, date_str), quantity in collected_data.items():
            # Extract the year and month from the date string
            year_month = date_str[:7]  # "YYYY-MM"
            key = (name, year_month)
            
            if key in monthly_sums:
                monthly_sums[key] += quantity
            else:
                monthly_sums[key] = quantity
        
        return monthly_sums

    def fill_table2(monthly_sums):
        # Extract the column indices for each month
        month_cols = {month: i for i, month in enumerate(header2[1:], 1)}

        # Fill the table with summed quantities
        for row in range(table2.rowCount()):
            name = table2.item(row, 0).text()
            for month, col_index in month_cols.items():
                year_month = month.replace("估量", "").replace("/", "-")
                key = (name, year_month)
                quantity = monthly_sums.get(key, 0)
                if (quantity != 0):
                    table2.setItem(row, col_index, QTableWidgetItem(str(quantity)))

    def handle_confirm_click(dialog, tab_widget, date):
        # Iterate through each tab
        for i in range(tab_widget.count()):
            tab = tab_widget.widget(i)
            tab_name = tab_widget.tabText(i)

            # Get the QTableWidget from the tab
            calendar_table = tab.findChild(QTableWidget)
            if not calendar_table:
                continue

            # Iterate through each cell in the calendar
            for row in range(calendar_table.rowCount()):
                for col in range(calendar_table.columnCount()):
                    cell_widget = calendar_table.cellWidget(row, col)
                    if not cell_widget:
                        continue

                    # Get the QSpinBox and QComboBox from the cell widget
                    spin_box = cell_widget.findChild(QSpinBox)

                    if spin_box:
                        quantity = spin_box.value()
                        if quantity > 0:
                            date_str = f"{date.strftime("%Y-%m")}-{cell_widget.findChild(QLabel).text().zfill(2)}"  # Example of date formatting, adjust if needed
                            collected_data.update({(tab_name, date_str):quantity})


        monthly_sums = sum_quantities_by_month()
        fill_table2(monthly_sums)
        # Close the dialog
        dialog.accept()

    def handle_header_click(logicalIndex):
        header_text = table2.horizontalHeaderItem(logicalIndex).text()

        # Create a dialog to show the calendar and quantity input
        dialog = QDialog()
        dialog.setWindowTitle(f"Quantity for {header_text}")  # Set the dialog title

        # Create a QVBoxLayout to hold the elements
        layout = QVBoxLayout(dialog)

        # Create a QLabel to show the header text
        label = QLabel(f"Enter quantities for: {header_text}", dialog)
        label.setFont(QFont('Arial', 14, QFont.Bold))
        layout.addWidget(label)

        # Create a QTabWidget to hold the tabs for each date
        tab_widget = QTabWidget()
        layout.addWidget(tab_widget)

        tab_bar = tab_widget.tabBar()
        font = QFont()
        font.setPointSize(17)  
        font.setBold(True)
        font.setFamily("Microsoft JhengHei")
        tab_bar.setFont(font)

        current_date = datetime.now()

        # Create a calendar for each date in first_column_data
        for date_str in first_column_data:
            tab = QWidget()
            tab_layout = QVBoxLayout(tab)
            tab.setLayout(tab_layout)

            # Create a QTableWidget to represent the calendar
            calendar_table = QTableWidget(6, 7, tab)  # 6 rows, 7 columns (like a calendar)
            calendar_table.setHorizontalHeaderLabels(['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'])
            calendar_table.verticalHeader().setVisible(False)
            calendar_table.horizontalHeader().setStretchLastSection(True)
            calendar_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            calendar_table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
            tab_layout.addWidget(calendar_table)

            # Convert the date_str to a datetime object
            try:
                date = datetime.strptime(header_text[:-2], "%Y/%m")
            except ValueError:
                continue  # Skip if date_str is not in the expected format

            # Determine the month and year for the calendar
            first_day_of_month = date.replace(day=1)
            start_day = first_day_of_month.weekday()+1  # Monday is 1, Sunday is 7
            days_in_month = (date.replace(month=date.month % 12 + 1, day=1) - first_day_of_month).days

            row, col = 0, start_day
            for day in range(1, days_in_month + 1):
                # Create a QWidget to hold the date, spin box, and combo box
                cell_widget = QWidget()
                cell_layout = QVBoxLayout()

                # Create a QLabel for the date
                date_label = QLabel(str(day))
                date_label.setAlignment(Qt.AlignCenter)
                date_label.setFont(QFont('Arial', 12))  # Increase font size

                # Create a QSpinBox for quantity input
                spin_box = QSpinBox()
                spin_box.setRange(0, 99999)
                spin_box.setAlignment(Qt.AlignCenter)
                spin_box.setFont(QFont('Arial', 10))  # Adjust font size
                spin_box.setFixedHeight(30)  # Increase the height of the spin box
                key = (date_str, f"{date.strftime("%Y-%m")}-{str(day).zfill(2)}")
                quantity = collected_data.get(key, 0)
                if (val != None):
                    spin_box.setValue(quantity)

                loop_date = datetime(date.year, date.month, day)

                if loop_date < current_date:
                    spin_box.setEnabled(False)

                # Add the date label, spin box, and combo box to the layout
                cell_layout.addWidget(date_label)
                cell_layout.addWidget(spin_box)
                cell_layout.setContentsMargins(5, 5, 5, 5)  # Add padding around the widgets
                cell_widget.setLayout(cell_layout)

                # Set the widget to the table cell
                calendar_table.setCellWidget(row, col % 7, cell_widget)

                col += 1
                if col % 7 == 0:
                    row += 1

            # Add the tab with the date_str as the title
            tab_widget.addTab(tab, date_str)

         # Create a Confirm button
        confirm_button = QPushButton("確認", dialog)
        confirm_button.setFont(QFont('Arial', 12, QFont.Bold))
        confirm_button.clicked.connect(lambda: handle_confirm_click(dialog, tab_widget, date))
        layout.addWidget(confirm_button)

        # Adjust size to fit the content
        dialog.setFixedSize(950, 700)

        # Show the dialog
        dialog.exec_()

    first_column_data = [table2.item(i, 0).text() for i in range(table2.rowCount())]

    set_header.sectionClicked.connect(handle_header_click)

    if isinstance(linedit_filename, QComboBox):
        monthly_sums = sum_quantities_by_month()
        fill_table2(monthly_sums)
        button_name = ','.join(table_list)
        button_bom.setText(button_name)

def read_demand(linedit_filename, combo_forecast, table):
    from main_page import user
    table.clear()
    
    # renew table --------------------------------------------------------------------------------

    if(linedit_filename.text() == "請點擊選擇BOM表" and combo_forecast.currentText() == ""):
        show_alert("請先選擇BOM表")
        return
    
    if combo_forecast.currentText() == "":
        table_list = linedit_filename.text().split(',')
    else:
        Bom_group_data = Name_Rule_SQL_handler.get_table_data(table_name="forecast_list", options={"EditBy":user}, database_name="forecast")
        table_list = list(set(item[2] for item in Bom_group_data))
        # print(table_list)

    table.setColumnCount(2+len(table_list))

    header = ["ERP Code"]

    text = ""

    for i in range(len(table_list)):
        text = text + (str(i+1) + "." +table_list[i].split('/')[-1].split('_')[0] + " ")
        header.append(table_list[i].split('/')[-1].split('_')[0]+"需量/1pc")

    header.append("庫存")

    table.setHorizontalHeaderLabels(header)

    erp_codes_list = []
    quantities_list = []
    final_list = []
    data = []

    for i in range(len(table_list)):
        df = SQL_handler.get_table_data(table_name=table_list[i], database_name='bom')
        for val in df:
            data.append({val[0]:val[1]})
        erp_codes_list.append([list(item.keys())[0] for item in data])
        quantities_list.append([list(item.values())[0] for item in data])
        final_list += [list(item.keys())[0] for item in data]

    none_repeat = []
    for item in final_list:
        if item not in none_repeat:
            none_repeat.append(item)

    table.setRowCount(len(none_repeat))

    combo_list = [none_repeat]

    for i in range(len(table_list)):
        combo_list.append([0] * len(none_repeat))

    for i in range(len(erp_codes_list)):
        for j, item in enumerate(none_repeat):
            if item in erp_codes_list[i]:
                index = erp_codes_list[i].index(item)
                combo_list[i + 1][j] = quantities_list[i][index]

    SQL_stock = SQL_handler.get_table_data(table_name='Inventory')
    

    # 初始化空字典
    SQL_stock_dict = {}

    # 遍歷資料
    for item in SQL_stock:
        key = item[1]
        quantity = item[2]
        
        if key in SQL_stock_dict:
            SQL_stock_dict[key] += quantity
        else:
            SQL_stock_dict[key] = quantity

    for row in range(len(combo_list[0])):
        for col in range(len(combo_list)):
            table.setItem(row, col, QTableWidgetItem(str(combo_list[col][row])))

        count = SQL_stock_dict.get(combo_list[0][row], 0)
        table.setItem(row, col+1, QTableWidgetItem(str(count)))
    # renew table end --------------------------------------------------------------------------------
    date_sums = {}

    def sum_quantities_by_date():
        # Iterate over the items in collected_data
        for (name, date_str), quantity in collected_data.items():
            # Use the full date string "YYYY-MM-DD" as the key
            if date_str in date_sums:
                org_quantity = date_sums[date_str]
                add_quantity = [(int(x) * quantity) for x in all_demand_data[name]]
                date_sums[date_str] = [x + y for x, y in zip(org_quantity, add_quantity)]
                # date_sums[date_str] += [x * quantity for x in all_demand_data[name]]
            else:
                date_sums[date_str] = [(int(x) * quantity) for x in all_demand_data[name]]   

    def get_column_data_by_header_name(header_name, back_size = -6):
        column_count = table.columnCount()

        column_index = -1
        for col in range(column_count):
            if back_size != 0:
                if header_name == table.horizontalHeaderItem(col).text()[:-6]: # 需量/1pcs
                    column_index = col
                    break
            else:
                if header_name == table.horizontalHeaderItem(col).text(): 
                    column_index = col
                    break
        
        if column_index == -1:
            return []

        column_data = []
        row_count = table.rowCount()
        for row in range(row_count):
            item = table.item(row, column_index)
            if item is not None:
                column_data.append(item.text())
            else:
                column_data.append('')

        return column_data
    
    all_demand_data = {}

    for col_name in table_list:
        all_demand_data.update({col_name:get_column_data_by_header_name(col_name)})

    # total_use_per_date = {} # {date: q_use}
    sum_quantities_by_date()

    # Extract and sort unique dates
    dates = sorted(set(date for _, date in collected_data.keys()))

    # Get the number of columns already in the table
    ColCount = 2+len(table_list)

    storage = get_column_data_by_header_name("庫存",0)

    for date in dates:
        # Insert a new column at the end
        table.insertColumn(ColCount)

        # Create and set the header item with the sorted date
        header_item = QTableWidgetItem(f"{date}缺料")
        table.setHorizontalHeaderItem(ColCount, header_item)

        storage = [int(x) - y for x, y in zip(storage, date_sums[date])]

        # Fill the new column with data from the storage list
        for row in range(table.rowCount()):
            # Set the data in the new column
            table.setItem(row, ColCount, QTableWidgetItem(str(storage[row])))

        # Update the column count 
        ColCount += 1

def combo_read_demand(combo_forecast, table):
    from main_page import user
    table.clear()

    # renew table --------------------------------------------------------------------------------

    if combo_forecast == "":
        show_alert("請先選擇BOM表")
        return
    
    Bom_group_data = Name_Rule_SQL_handler.get_table_data(table_name="forecast_list", options={"EditBy":user}, database_name="forecast")
    table_list = list(set(item[2] for item in Bom_group_data))
    print(table_list)

    table.setColumnCount(2+len(table_list))

    header = ["ERP Code"]

    text = ""

    for i in range(len(table_list)):
        text = text + (str(i+1) + "." +table_list[i].split('/')[-1].split('_')[0] + " ")
        header.append(table_list[i].split('/')[-1].split('_')[0]+"需量/1pc")

    header.append("庫存")

    table.setHorizontalHeaderLabels(header)

    erp_codes_list = []
    quantities_list = []
    final_list = []
    data = []

    for i in range(len(table_list)):
        df = SQL_handler.get_table_data(table_name=table_list[i], database_name='bom')
        for val in df:
            data.append({val[0]:val[1]})
        erp_codes_list.append([list(item.keys())[0] for item in data])
        quantities_list.append([list(item.values())[0] for item in data])
        final_list += [list(item.keys())[0] for item in data]

    none_repeat = []
    for item in final_list:
        if item not in none_repeat:
            none_repeat.append(item)

    table.setRowCount(len(none_repeat))

    combo_list = [none_repeat]

    for i in range(len(table_list)):
        combo_list.append([0] * len(none_repeat))

    for i in range(len(erp_codes_list)):
        for j, item in enumerate(none_repeat):
            if item in erp_codes_list[i]:
                index = erp_codes_list[i].index(item)
                combo_list[i + 1][j] = quantities_list[i][index]

    SQL_stock = SQL_handler.get_table_data(table_name='Inventory')
    

    # 初始化空字典
    SQL_stock_dict = {}

    # 遍歷資料
    for item in SQL_stock:
        key = item[1]
        quantity = item[2]
        
        if key in SQL_stock_dict:
            SQL_stock_dict[key] += quantity
        else:
            SQL_stock_dict[key] = quantity

    for row in range(len(combo_list[0])):
        for col in range(len(combo_list)):
            table.setItem(row, col, QTableWidgetItem(str(combo_list[col][row])))

        count = SQL_stock_dict.get(combo_list[0][row], 0)
        table.setItem(row, col+1, QTableWidgetItem(str(count)))
    # renew table end --------------------------------------------------------------------------------
    date_sums = {}

    def sum_quantities_by_date():
        # Iterate over the items in collected_data
        for (name, date_str), quantity in collected_data.items():
            # Use the full date string "YYYY-MM-DD" as the key
            if date_str in date_sums:
                org_quantity = date_sums[date_str]
                add_quantity = [(int(x) * quantity) for x in all_demand_data[name]]
                date_sums[date_str] = [x + y for x, y in zip(org_quantity, add_quantity)]
                # date_sums[date_str] += [x * quantity for x in all_demand_data[name]]
            else:
                date_sums[date_str] = [(int(x) * quantity) for x in all_demand_data[name]]   

    def get_column_data_by_header_name(header_name, back_size = -6):
        column_count = table.columnCount()

        column_index = -1
        for col in range(column_count):
            if back_size != 0:
                if header_name == table.horizontalHeaderItem(col).text()[:-6]: # 需量/1pcs
                    column_index = col
                    break
            else:
                if header_name == table.horizontalHeaderItem(col).text(): 
                    column_index = col
                    break
        
        if column_index == -1:
            return []

        column_data = []
        row_count = table.rowCount()
        for row in range(row_count):
            item = table.item(row, column_index)
            if item is not None:
                column_data.append(item.text())
            else:
                column_data.append('')

        return column_data
    
    all_demand_data = {}

    for col_name in table_list:
        all_demand_data.update({col_name:get_column_data_by_header_name(col_name)})

    # total_use_per_date = {} # {date: q_use}
    sum_quantities_by_date()

    # Extract and sort unique dates
    dates = sorted(set(date for _, date in collected_data.keys()))

    # Get the number of columns already in the table
    ColCount = 2+len(table_list)

    storage = get_column_data_by_header_name("庫存",0)

    for date in dates:
        # Insert a new column at the end
        table.insertColumn(ColCount)

        # Create and set the header item with the sorted date
        header_item = QTableWidgetItem(f"{date}缺料")
        table.setHorizontalHeaderItem(ColCount, header_item)

        storage = [int(x) - y for x, y in zip(storage, date_sums[date])]

        # Fill the new column with data from the storage list
        for row in range(table.rowCount()):
            # Set the data in the new column
            table.setItem(row, ColCount, QTableWidgetItem(str(storage[row])))

        # Update the column count 
        ColCount += 1

def load_excel():
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    _filePath, _ = QFileDialog.getOpenFileNames(None,"Open File", "", "Text Files (*.xlsx)", options=options)
    print(_filePath)

def del_bom_group(combo_forecast):
    if (combo_forecast.currentText() == ""):
        show_alert("請選擇要刪除的群組")
        return

    message = f"是否刪除群組{combo_forecast.currentText()}?"
    reply = QMessageBox.question(None, '確認刪除', message, 
                                 QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

    if reply == QMessageBox.Yes:
        forecast_name = combo_forecast.currentText()
        status = Name_Rule_SQL_handler.delete_by_conditions(table_name="forecast_list", conditions= {"Group_Name":forecast_name}, database_name="forecast")

        if status != True:
            show_alert(f"刪除群組時發生錯誤: {status}")
            return

    elif reply == QMessageBox.No:
        return

def BOM_import_page():

    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)

    #button widget
    linedit_filename = create_lineedit(0,0, 600)
    #button callback
    widgets["linedit_filename"].append(linedit_filename)

    button_back.clicked.connect(back_to_main_page)

def export_excel(table, table2):
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    file_path, _ = QFileDialog.getSaveFileName(None, "Save File", "shortage_list.xlsx", "Excel Files (*.xlsx);", options=options)
    if file_path:
        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'

        # Extract data from the first table
        rows = table.rowCount()
        columns = table.columnCount()
        data = []

        for row in range(rows):
            row_data = []
            for column in range(columns):
                item = table.item(row, column)
                row_data.append(item.text() if item else '')
                if (column == 0):
                    spec_data_list = Name_Rule_SQL_handler.get_table_data(table_name="material", options={"ERP_Code":item.text()},database_name="pcbmanagement")
                    if len(spec_data_list) != 0:
                        if spec_data_list[0][2] != None:
                            row_data.append(spec_data_list[0][2])
                        else:
                            row_data.append(spec_data_list[0][7])
                    else:
                        show_alert(f"資料庫中無ERP Code:{item.text()}")
                        row_data.append("")

            data.append(row_data)

        header = [table.horizontalHeaderItem(i).text() for i in range(columns)]
        header.insert(1,"規格")
        
        # Extract data from the second table
        rows2 = table2.rowCount()
        columns2 = table2.columnCount()

        data2 = []
        for row in range(rows2):
            row_data = []
            for column in range(columns2):
                item = table2.item(row, column)
                row_data.append(item.text() if item else '')
            data2.append(row_data)

        header2 = [table2.horizontalHeaderItem(i).text() for i in range(columns2)]

        data3 = []
        for (name, date_str), quantity in collected_data.items():
            data3.append([name, date_str, quantity])
        
        header3 = ["專案名稱", "日期", "數量"]

        print(header, header2, header3)

        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df1 = pd.DataFrame(data, columns=header)
                df2 = pd.DataFrame(data2, columns=header2)
                df3 = pd.DataFrame(data3, columns=header3)
                
                df1.to_excel(writer, sheet_name='Shotage_Sheet', index=False)
                df2.to_excel(writer, sheet_name='Forecast_Sheet_month', index=False)
                df3.to_excel(writer, sheet_name='Forecast_Sheet_date', index=False)
            
            show_alert(f"儲存{file_path}成功!!", "通知")
        except Exception as e:
            show_alert(f"儲存至{file_path}時發生錯誤: {str(e)}")

def Material_shortage_page():
    clear_widgets(main_page_widgets)
    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)

    #button widget
    button_import = create_button("匯入BOM", "#DDDDDD", 0, 0, width=150)
    button_import.setToolTip('選擇Excel文件匯入編碼以及數量進入資料庫')
    #button callback
    widgets["button_import"].append(button_import)

    #place global widgets on the grid
    grid.addWidget(widgets["button_import"][-1], 9, 6, 1, 2, alignment= Qt.AlignRight)
    

    # button widget
    button_export = create_button("匯出Excel", "#DDDDDD", 0, 0, width=150)
    # button callback
    widgets["button_export"].append(button_export)

    # place global widgets on the grid
    grid.addWidget(widgets["button_export"][-1], 9, 5, 1, 2, alignment= Qt.AlignCenter)

    #button widget
    linedit_filename = create_button("請點擊選擇BOM表","#DDDDDD",0,0, width=600)
    #button callback
    widgets["linedit_filename"].append(linedit_filename)

    #place global widgets on the grid
    grid.addWidget(widgets["linedit_filename"][-1], 0, 1, 1, 5)

    label_or = create_label("或",0,0,width=30)
    widgets['label_or'].append(label_or)

    grid.addWidget(widgets["label_or"][-1], 0, 4, 1, 1, alignment= Qt.AlignCenter)

    combo_forecast = create_combobox([],0,0,width=450)
    widgets['combo_forecast'].append(combo_forecast)
    combo_forecast.setEditable(True)
    combo_forecast.lineEdit().setPlaceholderText("選擇專案群組")

    grid.addWidget(widgets['combo_forecast'][-1], 0,4,1,4, alignment= Qt.AlignRight)

    table = create_table()
    widgets["table"].append(table)
    
    grid.addWidget(widgets["table"][-1], 4, 0, 5, 8)

    table2 = create_table()
    # table2.setFixedWidth(302)
    widgets["table2"].append(table2)

    grid.addWidget(widgets["table2"][-1], 1, 0, 2, 8)


    # button widget
    button_save_forecast = create_button("儲存專案群組", "#DDDDDD", 0, 0, width=150)
    # button callback
    widgets["button_save_forecast"].append(button_save_forecast)

    # place global widgets on the grid
    grid.addWidget(widgets["button_save_forecast"][-1], 3, 5, 1, 2, alignment= Qt.AlignCenter|Qt.AlignTop)

    button_delete_forecast = create_button("刪除專案群組", "#DDDDDD", 0, 0, width=150)
    widgets["button_delete_forecast"].append(button_delete_forecast)

    grid.addWidget(widgets["button_delete_forecast"][-1], 3, 6, 1, 2, alignment= Qt.AlignRight|Qt.AlignTop)
    # button_calculate = create_button("計算", "#DDDDDD", 0, 0, width=170)
    # widgets["button_calculate"].append(button_calculate)
    # grid.addWidget(widgets["button_calculate"][-1], 8, 5, 1, 2)

    button_hide_and_show = create_button("隱藏\n欄位", "rgba(221,221,221,0)", 0,0,width=29,height=35,font=10)
    widgets["button_hide_and_show"].append(button_hide_and_show)
    grid.addWidget(widgets["button_hide_and_show"][-1], 4, 0, 1, 2,alignment=Qt.AlignTop)

    
    button_back.clicked.connect(back_to_main_page)
    # button_hide_and_show.clicked.connect(lambda: read_demand(table, table2))
    # table.horizontalHeader().sectionClicked.connect(lambda: hideColumnOnClick(table, 2))
    # button_load_excel.clicked.connect(lambda: showDialog(table, linedit_filename, table2))  
    table2.itemChanged.connect(lambda: read_demand(linedit_filename, combo_forecast, table))
    button_hide_and_show.clicked.connect(lambda: hideColumnOnClick(table, table.columnCount(), button_hide_and_show))
    linedit_filename.clicked.connect(lambda: showDialog(linedit_filename, table, table2))
    button_export.clicked.connect(lambda: export_excel(table, table2))
    button_import.clicked.connect(import_excel)
    button_save_forecast.clicked.connect(lambda: save_bom_group(linedit_filename))
    button_delete_forecast.clicked.connect(lambda: del_bom_group(combo_forecast))
    def renew_combo_items():
        from main_page import user
        combo_forecast.blockSignals(True)
        Bom_group_data = Name_Rule_SQL_handler.get_table_data(table_name="forecast_list", options={"EditBy":user}, database_name="forecast")
        Bom_group_list = set(item[1] for item in Bom_group_data)
        combo_forecast.clear()
        combo_forecast.addItems(Bom_group_list)
        combo_forecast.setCurrentIndex(-1)
        combo_forecast.blockSignals(False)

    renew_combo_items()

    combo_forecast.currentIndexChanged.connect(lambda: table_show(combo_forecast, table, table2, linedit_filename))

    def save_bom_group(linedit_filename):
        from main_page import user

        if(linedit_filename.text() == "請點擊選擇BOM表"):
            show_alert("請先選擇BOM表")
            return

        table_list = linedit_filename.text().split(',')

        dialog = QDialog()
        dialog.setWindowTitle(f"儲存專案群組")  # Set the dialog title

        layout = QVBoxLayout(dialog)

        label = create_label("請輸入群組名稱",0,0)
        layout.addWidget(label)

        line_edit = create_lineedit(0,0)
        layout.addWidget(line_edit)

        button_layout = QHBoxLayout()

        # Create a QPushButton for saving
        save_button = create_button("確認儲存","#DDDDDD",0,0)
        button_layout.addWidget(save_button)

        group_data = []

        def click_button():

            for data in collected_data:
                date_obj = datetime.strptime(data[1], "%Y-%m-%d").date()
                group_data.append((line_edit.text(), data[0], collected_data[data], date_obj, user))

            status = SQL_handler.add_data_to_Forecast_Group_table(group_list=group_data, database_name="forecast")

            if status == True:
                show_alert(f"新增群組{line_edit.text()}成功","通知")
                renew_combo_items()
                dialog.accept()
            else:
                show_alert(f"新增群組{line_edit.text()}失敗: {status}")

        save_button.clicked.connect(click_button)

        layout.addLayout(button_layout)

        dialog.exec_()

        