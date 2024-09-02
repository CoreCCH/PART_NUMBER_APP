from PyQt5.QtWidgets import QLabel, QPushButton, QGraphicsDropShadowEffect
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt
from component import show_alert, grid,login_page_widgets, main_page_widgets, clear_widgets, generate_page_widgets, search_page_widgets, barcode_page_widgets, material_shortage_widgets, part_code_generate_page_widget, part_code_generate_page_hide_widget
from longin_page import user

def button_enter_event(shadow):
    shadow.setColor(QColor(50, 50 ,255, 255))

def main_page():
    clear_widgets(login_page_widgets)
    clear_widgets(main_page_widgets)
    clear_widgets(part_code_generate_page_widget)
    clear_widgets(part_code_generate_page_hide_widget)
    clear_widgets(search_page_widgets)
    clear_widgets(barcode_page_widgets)
    clear_widgets(material_shortage_widgets)

    title = QLabel("EPR輔助軟體")
    title.setFixedHeight(190)
    title.setStyleSheet('''
        *{
            font-size: 50px;
            
            color: 'black';
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        ''')
    #info widget
    main_page_widgets["title"].append(title)
    grid.addWidget(main_page_widgets["title"][-1], 1, 1, 3, 6)

    button1 = QPushButton("物料編碼建置")
    button1.setStyleSheet('''
        *{
            font-size: 30px;
            background: rgba(172,222,208,100);
            color: rgba(50,50,50,120);
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        *:hover{
            font-size: 30px;
            background: rgba(172,222,208,255);
            color: rgba(0,0,0,250);
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        ''')
    main_page_widgets["button1"].append(button1)
    
    from part_code_generator import part_code_generator_page
    button1.clicked.connect(part_code_generator_page)
    button1.setFixedSize(500,200)
    grid.addWidget(main_page_widgets["button1"][-1], 4, 1, 2, 3)

    button2 = QPushButton("物料查詢")
    button2.setStyleSheet('''
        *{
            font-size: 30px;
            background: rgba(172,222,208,100);
            color: rgba(50,50,50,120);
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        *:hover{
            font-size: 30px;
            background: rgba(172,222,208,255);
            color: rgba(0,0,0,250);
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        ''')
    main_page_widgets["button2"].append(button2)

    from part_code_searcher import frame_search_page
    button2.clicked.connect(frame_search_page)
    button2.setFixedSize(500,200)
    grid.addWidget(main_page_widgets["button2"][-1], 4, 4, 2, 3)

    button3 = QPushButton("物料入庫\n出庫\n退庫\b")
    button3.setStyleSheet('''
        *{
            font-size: 30px;
            background: rgba(172,222,208,100);
            color: rgba(50,50,50,120);
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        *:hover{
            font-size: 30px;
            background: rgba(172,222,208,255);
            color: rgba(0,0,0,250);
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        ''')
    # shadow_effect = QGraphicsDropShadowEffect()
    # shadow_effect.setBlurRadius(10)
    # shadow_effect.setXOffset(5)
    # shadow_effect.setYOffset(5)
    # shadow_effect.setColor(QColor(50, 50 ,50, 50))


    # button3.enterEvent(button_enter_event(shadow_effect))

    main_page_widgets["button3"].append(button3)
    from barcode_generator import barcode_select_page
    button3.clicked.connect(barcode_select_page)
    button3.setFixedSize(500,200)
    grid.addWidget(main_page_widgets["button3"][-1], 7, 1, 2, 3)

    button4 = QPushButton("欠料表")
    button4.setStyleSheet('''
        *{
            font-size: 30px;
            background: rgba(172,222,208,100);
            color: rgba(50,50,50,120);
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        *:hover{
            font-size: 30px;
            background: rgba(172,222,208,255);
            color: rgba(0,0,0,250);
            padding: 5px 10px;
            border-radius: 5px;
            font-family: Microsoft JhengHei;
            font-weight: bold;
        }
        ''')
    # shadow_effect = QGraphicsDropShadowEffect()
    # shadow_effect.setBlurRadius(10)
    # shadow_effect.setXOffset(5)
    # shadow_effect.setYOffset(5)
    # shadow_effect.setColor(QColor(50, 50 ,50, 50))


    # button3.enterEvent(button_enter_event(shadow_effect))

    main_page_widgets["button4"].append(button4)
    from Material_shortage_list import Material_shortage_page
    button4.clicked.connect(Material_shortage_page)
    button4.setFixedSize(500,200)
    grid.addWidget(main_page_widgets["button4"][-1], 7, 4, 2, 3)