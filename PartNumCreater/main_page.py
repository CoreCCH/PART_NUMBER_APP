from PyQt5.QtWidgets import QLabel, QPushButton
from component import grid, main_page_widgets, clear_widgets, generate_page_widgets, search_page_widgets

def main_page():
    clear_widgets(main_page_widgets)
    clear_widgets(generate_page_widgets)
    clear_widgets(search_page_widgets)

    title = QLabel("EPR輔助工具")
    #info widget
    main_page_widgets["title"].append(title)
    grid.addWidget(main_page_widgets["title"][-1], 1, 1, 3, 6)

    button1 = QPushButton("物料產生")
    main_page_widgets["button1"].append(button1)
    
    from part_code_generator import frame1
    button1.clicked.connect(frame1)
    button1.setFixedSize(200,300)
    grid.addWidget(main_page_widgets["button1"][-1], 4, 1, 3, 1)

    button2 = QPushButton("物料查詢")
    main_page_widgets["button2"].append(button2)
    from part_code_searcher import frame_search_page
    button2.clicked.connect(frame_search_page)
    button2.setFixedSize(200,300)
    grid.addWidget(main_page_widgets["button2"][-1], 4, 2, 3, 1)

    button3 = QPushButton("條碼產生")
    main_page_widgets["button3"].append(button3)
    button3.setFixedSize(200,300)
    grid.addWidget(main_page_widgets["button3"][-1], 4, 3, 3, 1)

    button4 = QPushButton("條碼讀取")
    main_page_widgets["button4"].append(button4)
    button4.setFixedSize(200,300)
    grid.addWidget(main_page_widgets["button4"][-1], 4, 4, 3, 1)