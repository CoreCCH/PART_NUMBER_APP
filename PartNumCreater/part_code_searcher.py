#GUI imports
import execl_handle
from component import grid, search_page_widgets as widgets, main_page_widgets, clear_widgets, create_button, create_lineedit, create_label, show_alert

# 定義函數根據part_number返回整列資訊
def get_part_info(df ,part_number):
    part_info = df[df['part_number'] == part_number]
    if not part_info.empty:
        return part_info.values.tolist()
    else:
        return None

def search_part(part_code: str, lineEdit1,lineEdit2, lineEdit3, lineEdit4, lineEdit5, lineEdit6, lineEdit7, lineEdit8, lineEdit9, lineEdit10, lineEdit11, lineEdit12):
    headers = ["part_number", "品項編號", "品項名稱", "項目", "種類", "尺寸/種類","%數", "容值/阻值/名稱", "電壓", "廠商", "供應商", "產生時間"]
    file_path = 'output.xlsx'

    if(len(part_code) != 11):
        show_alert("數入編碼長度錯誤")
        return
    
    output_df = execl_handle.check_output_existing(file_path, headers)
    if (get_part_info(output_df, part_code) == None):
        show_alert("零件編碼未建立")
    else:
        _list = get_part_info(output_df, part_code)[-1]
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

def frame_search_page():
    clear_widgets(main_page_widgets)
    clear_widgets(widgets)

    #button widget
    button_back = create_button("-", "#545454", 0, 0, 20, 60)
    #button callback
    widgets["button_back"].append(button_back)

    #place global widgets on the grid
    grid.addWidget(widgets["button_back"][-1], 0, 0, 1, 1)
    from main_page import main_page
    button_back.clicked.connect(main_page)

    #info widget
    label0 = create_label("料件編碼", 0, 0)
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

    lineEdit1 = create_lineedit(0, 0)
    widgets["line_bar1"].append(lineEdit1)
    grid.addWidget(widgets["line_bar1"][-1], 2, 2, 1, 2)

    #2nd
    label2 = create_label("編碼", 0, 0)
    widgets["label2"].append(label2)
    grid.addWidget(widgets["label2"][-1], 3, 1, 1, 1)

    lineEdit2 = create_lineedit(0,0)
    widgets["line_bar2"].append(lineEdit2)
    grid.addWidget(widgets["line_bar2"][-1], 3, 2, 1, 2)

    #3rd
    label3 = create_label("名稱", 0, 0)
    widgets["label3"].append(label3)
    grid.addWidget(widgets["label3"][-1], 4, 1, 1, 1)

    lineEdit3 = create_lineedit(0,0)
    widgets["line_bar3"].append(lineEdit3)
    grid.addWidget(widgets["line_bar3"][-1], 4, 2, 1, 2)

    #4th
    label4 = create_label("項目", 0, 0)
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 5, 1, 1, 1)

    lineEdit4 = create_lineedit(0,0)
    widgets["line_bar4"].append(lineEdit4)
    grid.addWidget(widgets["line_bar4"][-1], 5, 2, 1, 2)

    #5th
    label5 = create_label("種類", 0, 0)
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 6, 1, 1, 1)

    lineEdit5 = create_lineedit(0,0)
    widgets["line_bar5"].append(lineEdit5)
    grid.addWidget(widgets["line_bar5"][-1], 6, 2, 1, 2)

    #6th
    label6 = create_label("尺寸/種類", 0, 0)
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 7, 1, 1, 1)

    lineEdit6 = create_lineedit(0,0)
    widgets["line_bar6"].append(lineEdit6)
    grid.addWidget(widgets["line_bar6"][-1], 7, 2, 1, 2)

    #7th
    label7 = create_label("%數", 0, 0)
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 2, 4, 1, 1)

    lineEdit7 = create_lineedit(0,0)
    widgets["line_bar7"].append(lineEdit7)
    grid.addWidget(widgets["line_bar7"][-1], 2, 5, 1, 2)

    #8th
    label8 = create_label("容值/阻值/名稱", 0, 0)
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 3, 4, 1, 1)

    lineEdit8 = create_lineedit(0,0)
    widgets["line_bar8"].append(lineEdit8)
    grid.addWidget(widgets["line_bar8"][-1], 3, 5, 1, 2)
    
    #9th
    label9 = create_label("電壓", 0, 0)
    widgets["label9"].append(label9)
    grid.addWidget(widgets["label9"][-1], 4, 4, 1, 1)

    lineEdit9 = create_lineedit(0,0)
    widgets["line_bar9"].append(lineEdit9)
    grid.addWidget(widgets["line_bar9"][-1], 4, 5, 1, 2)

    #10th
    label10 = create_label("製造商", 0, 0)
    widgets["label10"].append(label10)
    grid.addWidget(widgets["label10"][-1], 5, 4, 1, 1)

    lineEdit10 = create_lineedit(0,0)
    widgets["line_bar10"].append(lineEdit10)
    grid.addWidget(widgets["line_bar10"][-1], 5, 5, 1, 2)

    #11th
    label11 = create_label("供應商", 0, 0)
    widgets["label11"].append(label11)
    grid.addWidget(widgets["label11"][-1], 6, 4, 1, 1)

    lineEdit11 = create_lineedit(0,0)
    widgets["line_bar11"].append(lineEdit11)
    grid.addWidget(widgets["line_bar11"][-1], 6, 5, 1, 2)

    #11th
    label12 = create_label("產生日期", 0, 0)
    widgets["label12"].append(label12)
    grid.addWidget(widgets["label12"][-1], 7, 4, 1, 1)

    lineEdit12 = create_lineedit(0,0)
    widgets["line_bar12"].append(lineEdit12)
    grid.addWidget(widgets["line_bar12"][-1], 7, 5, 1, 2)

    lineEdit1.setDisabled(True)
    lineEdit2.setDisabled(True)
    lineEdit3.setDisabled(True)
    lineEdit4.setDisabled(True)
    lineEdit5.setDisabled(True)
    lineEdit6.setDisabled(True)
    lineEdit7.setDisabled(True)
    lineEdit8.setDisabled(True)
    lineEdit9.setDisabled(True)
    lineEdit10.setDisabled(True)
    lineEdit11.setDisabled(True)
    lineEdit12.setDisabled(True)

    button1.clicked.connect(lambda: search_part(lineEdit0.text(), lineEdit1,lineEdit2, lineEdit3, lineEdit4, lineEdit5, lineEdit6, lineEdit7, lineEdit8, lineEdit9, lineEdit10, lineEdit11, lineEdit12))



