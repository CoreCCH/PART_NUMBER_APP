from component import grid, barcode_page_widgets as widgets, clear_widgets, main_page_widgets, create_button,create_label,create_lineedit, create_dateedit, create_combobox,show_alert,preview, barcode_reader
import execl_handle
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
}

stockplace = list(set(value[0] for value in stockroom.values()))
placecode = {"上海":0, "無錫":1, "深圳":2, "杭州":3, "竹北":4, "台南":5}

count = ["1","2","3","4","5","6","7","8","9","10"]

def update_combo2(__stockpalce, combo2):
    keys = [key for key, value in stockroom.items() if value[0] == __stockpalce]
    combo2.clear()
    combo2.addItems(keys)

def get_part_info(df ,part_number):
    part_info = df[df['part_number'] == part_number]
    if not part_info.empty:
        return part_info.values.tolist()
    else:
        return None

def sure_and_show(part_code, l1, l2, l3, l4, l5):
    headers = ["part_number", "品項編號", "品項名稱", "項目", "種類", "尺寸/種類","%數", "容值/阻值/名稱", "電壓", "廠商", "供應商", "產生時間"]
    file_path = 'output.xlsx'

    if(len(part_code) != 11):
        show_alert("數入編碼長度錯誤")
        return

    output_df = execl_handle.check_output_existing(file_path, headers)
    if (get_part_info(output_df, part_code) == None):
        l1.setVisible(False)
        l2.setVisible(False)
        l3.setVisible(False)
        l4.setVisible(False)
        l5.setVisible(False)
        show_alert("零件編碼未建立")
    else:
        l1.setVisible(True)
        l2.setVisible(True)
        l3.setVisible(True)
        l4.setVisible(True)
        l5.setVisible(True)

def button1_click():
    in_store_page()

def button2_click():
    out_store_page()

def button3_click():
    back_store_page()

def barcode_select_page():
    clear_widgets(main_page_widgets)

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


    button1.clicked.connect(button1_click)
    button2.clicked.connect(button2_click)
    button3.clicked.connect(button3_click)

def in_store_page():
    clear_widgets(widgets)

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

    label1 = create_label("料號", 0, 0, align='center')
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

    label4 = create_label("製表日期", 0, 0, align='center')
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 5, 1, 1, 1)

    date_choose2 = create_dateedit(0,0,width = 170)
    widgets["date_choose2"].append(date_choose2)
    grid.addWidget(widgets["date_choose2"][-1], 5, 2, 1, 2)

    label5 = create_label("樣品倉地區", 0, 0, align='center')
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 6, 1, 1, 1)

    combo1 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo1"].append(combo1)
    grid.addWidget(widgets["combo1"][-1], 6, 2, 1, 2)

    label6 = create_label("樣品倉", 0, 0, align='center')
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 7, 1, 1, 1)

    combo2 = create_combobox(stockplace,0,0,width=220,font_size=20)
    widgets["combo2"].append(combo2)
    grid.addWidget(widgets["combo2"][-1], 7, 2, 1, 2)

    button_sure = create_button("資料確認", "#008E8E", 0 ,0)
    widgets["button_sure"].append(button_sure)
    grid.addWidget(widgets["button_sure"][-1], 8, 1, 1, 1)

    label7 = create_label("列印張數", 0, 0, align='right')
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 2, 4, 1, 1)
    label7.setVisible(False)

    combo3 = create_combobox(count, 0, 0, width=140, font_size=20)
    widgets["combo3"].append(combo3)
    grid.addWidget(widgets["combo3"][-1], 2, 6, 1, 1)
    combo3.setEditable(True)
    combo3.setVisible(False)

    label8 = create_label("每筆數量", 0, 0, align='right')
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 3, 4, 1, 1)
    label8.setVisible(False)

    #LineEdit widget
    line_bar3 = create_lineedit(0, 0, width=140)
    widgets["line_bar3"].append(line_bar3)
    grid.addWidget(widgets["line_bar3"][-1], 3, 6, 1, 1)
    line_bar3.setVisible(False)

    button4 = create_button("預覽", "#DDDDDD", 0, 0)
    widgets["button4"].append(button4)
    button4.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["button4"][-1], 4, 6, 1, 1)


    combo1.currentIndexChanged.connect(lambda: update_combo2(combo1.currentText(),combo2))
    button2.clicked.connect(button2_click)
    button3.clicked.connect(button3_click)
    button_sure.clicked.connect(lambda: sure_and_show(line_bar1.text(), label7, label8, combo3, line_bar3, button4))
    button4.clicked.connect(lambda: preview(int(combo3.currentText()), line_bar1.text(), int(line_bar3.text()), date_choose1.date().toString("yyyy-MM-dd"), date_choose2.date().toString("yyyy-MM-dd"), combo2.currentText(), int(line_bar2.text()), 0))

def out_store_page():
    clear_widgets(widgets)

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

    # label1 = create_label("料號", 0, 0, align='center')
    # widgets["label1"].append(label1)
    # grid.addWidget(widgets["label1"][-1], 2, 1, 1, 1)

    button_input = create_button("條碼機", "#DDDDDD", 0 ,0)
    widgets["button_input"].append(button_input)
    grid.addWidget(widgets["button_input"][-1], 2, 1, 1, 1)

    #LineEdit widget
    line_bar1 = create_lineedit(0,0,width=170)
    widgets["line_bar1"].append(line_bar1)
    grid.addWidget(widgets["line_bar1"][-1], 2, 2, 1, 2)

    label2 = create_label("數量", 0, 0, align='center')
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

    label4 = create_label("製表日期", 0, 0, align='center')
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 5, 1, 1, 1)

    date_choose2 = create_dateedit(0,0,width = 170)
    widgets["date_choose2"].append(date_choose2)
    grid.addWidget(widgets["date_choose2"][-1], 5, 2, 1, 2)

    label5 = create_label("樣品倉地區", 0, 0, align='center')
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 6, 1, 1, 1)

    combo1 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo1"].append(combo1)
    grid.addWidget(widgets["combo1"][-1], 6, 2, 1, 2)

    label6 = create_label("樣品倉", 0, 0, align='center')
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 7, 1, 1, 1)

    combo2 = create_combobox(stockplace,0,0,width=220,font_size=20)
    widgets["combo2"].append(combo2)
    grid.addWidget(widgets["combo2"][-1], 7, 2, 1, 2)

    button_sure = create_button("資料確認", "#008E8E", 0 ,0)
    widgets["button_sure"].append(button_sure)
    grid.addWidget(widgets["button_sure"][-1], 8, 1, 1, 1)

    label7 = create_label("列印張數", 0, 0, align='right')
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 2, 4, 1, 1)
    label7.setVisible(False)

    combo3 = create_combobox(count, 0, 0, width=140, font_size=20)
    combo3.setCurrentIndex(0)
    widgets["combo3"].append(combo3)
    grid.addWidget(widgets["combo3"][-1], 2, 6, 1, 1)
    combo3.setEditable(False)
    combo3.setDisabled(True)
    combo3.setVisible(False)

    label8 = create_label("出庫數量", 0, 0, align='right')
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 3, 4, 1, 1)
    label8.setVisible(False)

    #LineEdit widget
    line_bar3 = create_lineedit(0, 0, width=140)
    widgets["line_bar3"].append(line_bar3)
    grid.addWidget(widgets["line_bar3"][-1], 3, 6, 1, 1)
    line_bar3.setVisible(False)

    button4 = create_button("預覽", "#DDDDDD", 0, 0)
    widgets["button4"].append(button4)
    button4.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["button4"][-1], 4, 6, 1, 1)

    combo1.currentIndexChanged.connect(lambda: update_combo2(combo1.currentText(),combo2))
    button1.clicked.connect(button1_click)
    button3.clicked.connect(button3_click)
    button_sure.clicked.connect(lambda: sure_and_show(line_bar1.text(), label7, label8, combo3, line_bar3, button4))
    button_input.clicked.connect(lambda: barcode_reader(line_bar1, line_bar2, date_choose1, combo1, combo2))
    button4.clicked.connect(lambda: preview(int(combo3.currentText()), line_bar1.text(), int(line_bar3.text()), date_choose1.date().toString("yyyy-MM-dd"), date_choose2.date().toString("yyyy-MM-dd"), combo2.currentText(), int(line_bar2.text()), 1))

def back_store_page():
    clear_widgets(widgets)

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

    button_input = create_button("條碼機", "#DDDDDD", 0 ,0)
    widgets["button_input"].append(button_input)
    grid.addWidget(widgets["button_input"][-1], 2, 1, 1, 1)

    #LineEdit widget
    line_bar1 = create_lineedit(0,0,width=170)
    widgets["line_bar1"].append(line_bar1)
    grid.addWidget(widgets["line_bar1"][-1], 2, 2, 1, 2)

    label2 = create_label("數量", 0, 0, align='center')
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

    label4 = create_label("製表日期", 0, 0, align='center')
    widgets["label4"].append(label4)
    grid.addWidget(widgets["label4"][-1], 5, 1, 1, 1)

    date_choose2 = create_dateedit(0,0,width = 170)
    widgets["date_choose2"].append(date_choose2)
    grid.addWidget(widgets["date_choose2"][-1], 5, 2, 1, 2)

    label5 = create_label("樣品倉地區", 0, 0, align='center')
    widgets["label5"].append(label5)
    grid.addWidget(widgets["label5"][-1], 6, 1, 1, 1)

    combo1 = create_combobox(stockplace,0,0,width=170,font_size=20)
    widgets["combo1"].append(combo1)
    grid.addWidget(widgets["combo1"][-1], 6, 2, 1, 2)

    label6 = create_label("樣品倉", 0, 0, align='center')
    widgets["label6"].append(label6)
    grid.addWidget(widgets["label6"][-1], 7, 1, 1, 1)

    combo2 = create_combobox(stockplace,0,0,width=220,font_size=20)
    widgets["combo2"].append(combo2)
    grid.addWidget(widgets["combo2"][-1], 7, 2, 1, 2)

    button_sure = create_button("資料確認", "#008E8E", 0 ,0)
    widgets["button_sure"].append(button_sure)
    grid.addWidget(widgets["button_sure"][-1], 8, 1, 1, 1)

    label7 = create_label("列印張數", 0, 0, align='right')
    widgets["label7"].append(label7)
    grid.addWidget(widgets["label7"][-1], 2, 4, 1, 1)
    label7.setVisible(False)

    combo3 = create_combobox(count, 0, 0, width=140, font_size=20)
    combo3.setCurrentIndex(0)
    widgets["combo3"].append(combo3)
    grid.addWidget(widgets["combo3"][-1], 2, 6, 1, 1)
    combo3.setEditable(False)
    combo3.setDisabled(True)
    combo3.setVisible(False)

    label8 = create_label("出庫數量", 0, 0, align='right')
    widgets["label8"].append(label8)
    grid.addWidget(widgets["label8"][-1], 3, 4, 1, 1)
    label8.setVisible(False)

    #LineEdit widget
    line_bar3 = create_lineedit(0, 0, width=140)
    widgets["line_bar3"].append(line_bar3)
    grid.addWidget(widgets["line_bar3"][-1], 3, 6, 1, 1)
    line_bar3.setVisible(False)

    button4 = create_button("預覽", "#DDDDDD", 0, 0)
    widgets["button4"].append(button4)
    button4.setVisible(False)

    #place global widgets on the grid
    grid.addWidget(widgets["button4"][-1], 4, 6, 1, 1)

    combo1.currentIndexChanged.connect(lambda: update_combo2(combo1.currentText(),combo2))
    button1.clicked.connect(button1_click)
    button2.clicked.connect(button2_click)
    button_sure.clicked.connect(lambda: sure_and_show(line_bar1.text(), label7, label8, combo3, line_bar3, button4))
    button_input.clicked.connect(lambda: barcode_reader(line_bar1, line_bar2, date_choose1, combo1, combo2))
    button4.clicked.connect(lambda: preview(int(combo3.currentText()), line_bar1.text(), int(line_bar3.text()), date_choose1.date().toString("yyyy-MM-dd"), date_choose2.date().toString("yyyy-MM-dd"), combo2.currentText(), int(line_bar2.text()), 2))
