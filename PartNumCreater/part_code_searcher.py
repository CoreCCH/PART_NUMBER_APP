#GUI imports
import execl_handle
import SQL_handler
from component import grid, search_page_widgets as widgets, main_page_widgets, clear_widgets, create_button, create_lineedit, create_label, show_alert, barcode_reader_for_search, search_material
from PyQt5.QtCore import Qt

# 定義函數根據part_number返回整列資訊
def get_part_info(df ,part_number):
    part_info = df[df['ERP Code'] == part_number]
    if not part_info.empty:
        return part_info.values.tolist()
    else:
        return None

def search_part(part_code: str, lineEdit1,lineEdit2, lineEdit3, lineEdit4, lineEdit5, lineEdit6, lineEdit7, lineEdit8, lineEdit9, lineEdit10, lineEdit11, lineEdit12, lineEdit_pn, lineEdit13, lineEdit14, lineEdit15, lineEdit16, lineEdit17):
    
    # file_path = 'output.xlsx'

    if(len(part_code) != 11 and len(part_code) != 20):
        show_alert("數入編碼長度錯誤")
        return
    
    # output_df = execl_handle.check_output_existing(file_path)
    SQL_output = SQL_handler.fetch_data_from_Material_table({"ERP_Code":part_code[0:11]})

    if isinstance(SQL_output, str):
        show_alert(f"操作資料庫時發生錯誤: {SQL_output}")
    elif (SQL_output == []):
        show_alert("零件編碼未建立")
    else:
        _list = SQL_output[0]
        lineEdit1.setText(str(_list[1]))
        lineEdit2.setText(str(_list[0]))
        lineEdit3.setText(str(_list[2]))
        lineEdit4.setText(str(_list[3]))
        lineEdit5.setText(str(_list[4]))
        lineEdit6.setText(str(_list[5]))
        lineEdit7.setText(str(_list[6]))
        lineEdit8.setText(str(_list[7]))
        lineEdit9.setText(str(_list[8]))
        lineEdit10.setText(str(_list[9]))
        lineEdit11.setText(str(_list[10]))
        lineEdit12.setText(str(_list[11]))
        lineEdit_pn.setText(str(_list[12]))

    if len(part_code) > 11:
        data = barcode_reader_for_search(part_code)
        if data == False:
            return
        else:
            lineEdit13.setText(data[0])
            lineEdit14.setText(data[3])
            lineEdit15.setText(data[1])
            lineEdit16.setText(data[4])
            lineEdit17.setText(data[2])

def Input_info_to_search_material(line_materialID, line_Ecount, line_category, line_type, line_size, line_percentage, line_name, line_volt, line_manu, line_sup, line_createAt, line_pn):
    filter = {}
    if line_materialID.text() != "" and line_materialID.text() != 'None':
        # print(f"line_materialID: {line_materialID.text()}")
        filter.update({'ERP_Code':line_materialID.text()})
    if line_Ecount.text() != "" and line_Ecount.text() != 'None':
        # print(f"line_Ecount: {line_Ecount.text()}")
        filter.update({'ECount_Code':line_Ecount.text()})
    if line_category.text() != "" and line_category.text() != 'None':
        # print(f"line_category: {line_category.text()}")
        filter.update({'Item':line_category.text()})
    if line_type.text() != "" and line_type.text() != 'None':
        # print(f"line_type: {line_type.text()}")
        filter.update({'Type':line_type.text()})
    if line_size.text() != "" and line_size.text() != 'None':
        # print(f"line_size: {line_size.text()}")
        filter.update({'Size_Type':line_type.text()})
    if line_percentage.text() != "" and line_percentage.text() != 'None':
        # print(f"line_percentage: {line_percentage.text()}")
        filter.update({'Percentage_Package':line_percentage.text()})
    if line_name.text() != "" and line_name.text() != 'None':
        # print(f"line_name: {line_name.text()}")
        filter.update({'Capacitance_Resistance_Name':line_name.text()})
    if line_volt.text() != "" and line_volt.text() != 'None':
        # print(f"line_volt: {line_volt.text()}")
        filter.update({'Voltage_PinSize_Frequency':line_volt.text()})
    if line_manu.text() != "" and line_manu.text() != 'None':
        # print(f"line_manu: {line_manu.text()}")
        filter.update({'Manufacturer':line_manu.text()})
    if line_sup.text() != "" and line_sup.text() != 'None':
        # print(f"line_sup: {line_sup.text()}")
        filter.update({'Supplier':line_sup.text()})
    if line_createAt.text() != "" and line_createAt.text() != 'None':
        # print(f"line_createAt: {line_createAt.text()}")
        filter.update({'CreatedAt':line_createAt.text()})
    if line_pn.text() != "" and line_pn.text() != 'None':
        # print(f"line_pn: {line_pn.text()}")
        filter.update({'PartNumber':line_pn.text()})

    search_material(filter)


def frame_search_page():
    clear_widgets(main_page_widgets)
    clear_widgets(widgets)

    #button widget
    button_back = create_button("←", "#DDDDDD", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    from main_page import main_page
    button_back.clicked.connect(main_page)

    #info widget
    label0 = create_label("BarCode", 0, 0)
    widgets["label0"].append(label0)
    grid.addWidget(widgets["label0"][-1], 1, 1, 1, 1)

    #LineEdit widget
    lineEdit0 = create_lineedit(0,0)
    widgets["line_bar0"].append(lineEdit0)
    grid.addWidget(widgets["line_bar0"][-1], 1, 2, 1, 2)

    # widgets["logo"].append(logo)

    #button widget
    button1 = create_button("輸入", "#008E8E", 0, 0)
    #button callback
    # button.clicked.connect(start_game)
    widgets["button"].append(button1)

    #place global widgets on the grid
    grid.addWidget(widgets["button"][-1], 1, 4, 1, 1)

    #1st
    label1 = create_label("編號", 0, 0)
    widgets["label1"].append(label1)
    grid.addWidget(widgets["label1"][-1], 2, 1, 1, 1)

    lineEdit1 = create_lineedit(0, 0, 200)
    widgets["line_bar1"].append(lineEdit1)
    grid.addWidget(widgets["line_bar1"][-1], 2, 1, 1, 2, alignment=Qt.AlignRight)

    #2nd
    label2 = create_label("編碼", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 3, 1, 1, 1)

    lineEdit2 = create_lineedit(0,0,200)
    widgets["line_bar2"].append(lineEdit2)
    grid.addWidget(widgets["line_bar2"][-1], 3, 1, 1, 2, alignment=Qt.AlignRight)

    #3rd
    label3 = create_label("品項名稱", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 4, 1, 1, 1)

    lineEdit3 = create_lineedit(0,0,200)
    widgets["line_bar3"].append(lineEdit3)
    grid.addWidget(widgets["line_bar3"][-1], 4, 1, 1, 2, alignment=Qt.AlignRight)

    #4th
    label4 = create_label("項目", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 5, 1, 1, 1)

    lineEdit4 = create_lineedit(0,0,200)
    widgets["line_bar4"].append(lineEdit4)
    grid.addWidget(widgets["line_bar4"][-1], 5, 1, 1, 2, alignment=Qt.AlignRight)

    #5th
    label5 = create_label("種類", 0, 0)
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 6, 1, 1, 1)

    lineEdit5 = create_lineedit(0,0,200)
    widgets["line_bar5"].append(lineEdit5)
    grid.addWidget(widgets["line_bar5"][-1], 6, 1, 1, 2, alignment=Qt.AlignRight)

    #6th
    label6 = create_label("尺寸/種類", 0, 0)
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 7, 1, 1, 1)

    lineEdit6 = create_lineedit(0,0,200)
    widgets["line_bar6"].append(lineEdit6)
    grid.addWidget(widgets["line_bar6"][-1], 7, 1, 1, 2, alignment=Qt.AlignRight)

    #7th
    label7 = create_label("%數", 0, 0)
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 2, 3, 1, 1)

    lineEdit7 = create_lineedit(0,0,200)
    widgets["line_bar7"].append(lineEdit7)
    grid.addWidget(widgets["line_bar7"][-1], 2, 3, 1, 2, alignment=Qt.AlignRight)

    #8th
    label8 = create_label("容值/阻值/名稱", 0, 0)
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 3, 3, 1, 1)

    lineEdit8 = create_lineedit(0,0,200)
    widgets["line_bar8"].append(lineEdit8)
    grid.addWidget(widgets["line_bar8"][-1], 3, 3, 1, 2, alignment=Qt.AlignRight)
    
    #9th
    label9 = create_label("電壓", 0, 0)
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 4, 3, 1, 1)

    lineEdit9 = create_lineedit(0,0,200)
    widgets["line_bar9"].append(lineEdit9)
    grid.addWidget(widgets["line_bar9"][-1], 4, 3, 1, 2, alignment=Qt.AlignRight)

    #10th
    label10 = create_label("製造商", 0, 0)
    widgets["label10"].append(label10)
    grid.addWidget(widgets["label10"][-1], 5, 3, 1, 1)

    lineEdit10 = create_lineedit(0,0,200)
    widgets["line_bar10"].append(lineEdit10)
    grid.addWidget(widgets["line_bar10"][-1], 5, 3, 1, 2, alignment=Qt.AlignRight)

    #11th
    label11 = create_label("供應商", 0, 0)
    widgets["label11"].append(label11)
    grid.addWidget(widgets["label11"][-1], 6, 3, 1, 1)

    lineEdit11 = create_lineedit(0,0,200)
    widgets["line_bar11"].append(lineEdit11)
    grid.addWidget(widgets["line_bar11"][-1], 6, 3, 1, 2, alignment=Qt.AlignRight)

    #11th
    label12 = create_label("產生日期", 0, 0)
    widgets["label12"].append(label12)
    grid.addWidget(widgets["label12"][-1], 7, 3, 1, 1)

    lineEdit12 = create_lineedit(0,0,200)
    widgets["line_bar12"].append(lineEdit12)
    grid.addWidget(widgets["line_bar12"][-1], 7, 3, 1, 2, alignment=Qt.AlignRight)

    #button search
    button_search = create_button("搜尋","#DDDDDD",0,0)
    widgets["button_search"].append(button_search)
    grid.addWidget(widgets["button_search"][-1], 8, 3, 1, 2, alignment=Qt.AlignRight)

    #11th
    label_pn = create_label("PN", 0, 0)
    widgets["label_pn"].append(label_pn)
    grid.addWidget(widgets["label_pn"][-1], 8, 1, 1, 1)

    lineEdit_pn = create_lineedit(0,0,250)
    widgets["line_bar_pn"].append(lineEdit_pn)
    grid.addWidget(widgets["line_bar_pn"][-1], 8, 1, 1, 2, alignment=Qt.AlignRight)

    label13 = create_label("入庫日期", 0, 0, width=100)
    widgets["label13"].append(label13)
    grid.addWidget(widgets["label13"][-1], 2, 5, 1, 2, alignment=Qt.AlignLeft)

    lineEdit13 = create_lineedit(0,0,250)
    widgets["line_bar13"].append(lineEdit13)
    grid.addWidget(widgets["line_bar13"][-1], 2, 6, 1, 2, alignment=Qt.AlignRight)
    lineEdit13.setEnabled(False)

    label14 = create_label("數量", 0, 0, width=100)
    widgets["label14"].append(label14)
    grid.addWidget(widgets["label14"][-1], 3, 5, 1, 2, alignment=Qt.AlignLeft)

    lineEdit14 = create_lineedit(0,0,250)
    widgets["line_bar14"].append(lineEdit14)
    grid.addWidget(widgets["line_bar14"][-1], 3, 6, 1, 2, alignment=Qt.AlignRight)
    lineEdit14.setEnabled(False)

    label15 = create_label("原始庫", 0, 0, width=100)
    widgets["label15"].append(label15)
    grid.addWidget(widgets["label15"][-1], 4, 5, 1, 2, alignment=Qt.AlignLeft)

    lineEdit15 = create_lineedit(0,0,250)
    widgets["line_bar15"].append(lineEdit15)
    grid.addWidget(widgets["line_bar15"][-1], 4, 6, 1, 2, alignment=Qt.AlignRight)
    lineEdit15.setEnabled(False)



    label16 = create_label("箱號", 0, 0, width=100)
    widgets["label16"].append(label16)
    grid.addWidget(widgets["label16"][-1], 5, 5, 1, 2, alignment=Qt.AlignLeft)

    lineEdit16 = create_lineedit(0,0,250)
    widgets["line_bar16"].append(lineEdit16)
    grid.addWidget(widgets["line_bar16"][-1], 5, 6, 1, 2, alignment=Qt.AlignRight)
    lineEdit16.setEnabled(False)

    label17 = create_label("現庫", 0, 0, width=100)
    widgets["label17"].append(label17)
    grid.addWidget(widgets["label17"][-1], 6, 5, 1, 2, alignment=Qt.AlignLeft)

    lineEdit17 = create_lineedit(0,0,250)
    widgets["line_bar17"].append(lineEdit17)
    grid.addWidget(widgets["line_bar17"][-1], 6, 6, 1, 2, alignment=Qt.AlignRight)
    lineEdit17.setEnabled(False)

    # lineEdit1.setDisabled(True)
    # lineEdit2.setDisabled(True)
    # lineEdit3.setDisabled(True)
    # lineEdit4.setDisabled(True)
    # lineEdit5.setDisabled(True)
    # lineEdit6.setDisabled(True)
    # lineEdit7.setDisabled(True)
    # lineEdit8.setDisabled(True)
    # lineEdit9.setDisabled(True)
    # lineEdit10.setDisabled(True)
    # lineEdit11.setDisabled(True)
    # lineEdit12.setDisabled(True)
    # lineEdit_pn.setDisabled(True)

    button1.clicked.connect(lambda: search_part(lineEdit0.text(), lineEdit1,lineEdit2, lineEdit3, lineEdit4, lineEdit5, lineEdit6, lineEdit7, lineEdit8, lineEdit9, lineEdit10, lineEdit11, lineEdit12, lineEdit_pn, lineEdit13, lineEdit14, lineEdit15, lineEdit16, lineEdit17))
    button_search.clicked.connect(lambda: Input_info_to_search_material(lineEdit2, lineEdit1, lineEdit4, lineEdit5, lineEdit6, lineEdit7, lineEdit8, lineEdit9, lineEdit10, lineEdit11, lineEdit12, lineEdit_pn))



