from res import main_grid, main_widget, create_label, create_button, clear_widgets
from PyQt5.QtWidgets import QVBoxLayout, QPushButton, QApplication
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt, QSize, QUrl
from component import grid
from part_code_generator import part_code_generator_page
# from PyQt5.QtWebEngineWidgets import QWebEngineView
import sys

def main_page():
    clear_widgets(main_widget)

    nav_bar = create_label(name="", l_margin=0, r_margin=0, font_color="rgba(255,255,255,255)", font_size=20, 
                           background_color="rgba(238, 231, 218, 255)", align= "center")
    main_widget['nav-bar'].append(nav_bar)
    main_grid.addWidget(main_widget['nav-bar'][-1], 0,0,40,8)

    main_label = create_label(name="", l_margin=0, r_margin=0, font_color="rgba(255,255,255,255)", 
                              font_size=20, background_color="rgba(242, 241, 235, 255)", align= "center")
    main_widget['main-label'].append(main_label)
    main_grid.addWidget(main_widget['main-label'][-1], 0,8,40,42)

    # # show OST website??
    # browser_layout = QVBoxLayout()
    # main_label.setLayout(browser_layout)
    # browser = QWebEngineView()
    # url = QUrl("https://www.orient-suntech.com/")
    # browser.setUrl(url)
    # browser_layout.addWidget(browser)
    main_label.setLayout(grid)
    part_code_generator_page()

    diagonally_button = create_button(name="", l_margin=0, r_margin=0, font_color= "rgba(0,0,0,255)"
                                    , font_size=20, background_color="rgba(238, 231, 218, 255)"
                                    , hover_background_color= "rgba(238, 231, 218, 255)"
                                    , hover_font_color= "rgba(255,255,255,255)", hover_font_size= 20
                                    , hover_pic= "img/diagonally-arrow_.png")
    main_widget['diagonally-button'].append(diagonally_button)
    diagonally_button.setIcon(QIcon('img/diagonally-arrow.png'))
    diagonally_button.setIconSize(QSize(40, 40))
    main_grid.addWidget(main_widget['diagonally-button'][-1], 0,0,2,2)

    ost_logo_label = create_label(name="", l_margin=0, r_margin=0, font_color="rgba(255,255,255,255)", 
                                  font_size=20, background_color="rgba(238, 231, 218, 255)", align= "center")
    main_widget['ost-logo-label'].append(ost_logo_label)
    main_grid.addWidget(main_widget['ost-logo-label'][-1], 3,0,3,8)

    pixmap = QPixmap('img/OST_LOGO.png')
    scaled_pixmap = pixmap.scaled(QSize(150,280), Qt.KeepAspectRatio, Qt.SmoothTransformation)
    ost_logo_label.setPixmap(scaled_pixmap)

    ERP_code_button = create_button(name="物料編碼", l_margin=0, r_margin=0, font_color= "rgba(0,0,0,255)"
                                    , font_size=20, background_color="rgba(238, 231, 218, 255)"
                                    , hover_background_color= "rgba(197, 112, 93, 255)"
                                    , hover_font_color= "rgba(255,255,255,255)", hover_font_size= 22
                                    , hover_pic= "img/motherboard_.png")
    main_widget['ERP-code-button'].append(ERP_code_button)
    ERP_code_button.setIcon(QIcon('img/motherboard.png'))
    ERP_code_button.setIconSize(QSize(50, 50))
    main_grid.addWidget(main_widget['ERP-code-button'][-1], 6,0,4,7)
    
    search_button = create_button(name="搜尋物料", l_margin=0, r_margin=0, font_color= "rgba(0,0,0,255)"
                                    , font_size=20, background_color="rgba(238, 231, 218, 255)"
                                    , hover_background_color= "rgba(197, 112, 93, 255)"
                                    , hover_font_color= "rgba(255,255,255,255)", hover_font_size= 22
                                    , hover_pic= "img/magnifier_.png")
    main_widget['search-button'].append(search_button)
    search_button.setIcon(QIcon('img/magnifier.png'))
    search_button.setIconSize(QSize(50, 50))
    main_grid.addWidget(main_widget['search-button'][-1], 10,0,4,7)

    stock_button = create_button(name="物料管理", l_margin=0, r_margin=0, font_color= "rgba(0,0,0,255)"
                                    , font_size=20, background_color="rgba(238, 231, 218, 255)"
                                    , hover_background_color= "rgba(197, 112, 93, 255)"
                                    , hover_font_color= "rgba(255,255,255,255)", hover_font_size= 22
                                    , hover_pic= "img/warehouse_.png")
    main_widget['stock-button'].append(stock_button)
    stock_button.setIcon(QIcon('img/warehouse.png'))
    stock_button.setIconSize(QSize(50, 50))
    main_grid.addWidget(main_widget['stock-button'][-1], 14,0,4,7)


    mrp_button = create_button(name="需求規劃", l_margin=0, r_margin=0, font_color= "rgba(0,0,0,255)"
                                    , font_size=20, background_color="rgba(238, 231, 218, 255)"
                                    , hover_background_color= "rgba(197, 112, 93, 255)"
                                    , hover_font_color= "rgba(255,255,255,255)", hover_font_size= 22
                                    , hover_pic= "img/purchase-list_.png")
    main_widget['MRP-button'].append(mrp_button)
    mrp_button.setIcon(QIcon('img/purchase-list.png'))
    mrp_button.setIconSize(QSize(50, 50))
    main_grid.addWidget(main_widget['MRP-button'][-1], 18,0,4,7)

    ERP_code_layout = QVBoxLayout()
    main_widget['ERP-code-layout'].append(ERP_code_layout)
    
    ERP_code_generate_button = create_button(name="編碼建立", l_margin=0, r_margin=0, font_color= "rgba(0,0,0,255)"
                                            , font_size=20, background_color="rgba(238, 231, 218, 255)"
                                            , hover_background_color= "rgba(197, 112, 93, 255)"
                                            , hover_font_color= "rgba(255,255,255,255)", hover_font_size= 22)
    main_widget['ERP-code-generate-button'].append(ERP_code_generate_button)
    ERP_code_layout.addWidget(main_widget['ERP-code-generate-button'][-1])

    rule_generate_button = create_button(name="新增零件規則", l_margin=0, r_margin=0, font_color= "rgba(0,0,0,255)"
                                            , font_size=20, background_color="rgba(238, 231, 218, 255)"
                                            , hover_background_color= "rgba(197, 112, 93, 255)"
                                            , hover_font_color= "rgba(255,255,255,255)", hover_font_size= 22)
    main_widget['rule-generate-button'].append(rule_generate_button)
    ERP_code_layout.addWidget(main_widget['rule-generate-button'][-1])
    main_grid.addLayout(ERP_code_layout, 5, 7, 4, 7)
    hide_sub_layout(ERP_code_layout)

    ERP_code_button.clicked.connect(lambda: toggle_sub_layout(ERP_code_layout))
    diagonally_button.clicked.connect(reduce_main_page)
    
def reduce_main_page():
    main_grid.addWidget(main_widget['nav-bar'][-1], 0,0,40,3)
    main_grid.addWidget(main_widget['main-label'][-1], 0,3,40,47)

    main_widget['diagonally-button'][-1].setText('')
    main_grid.addWidget(main_widget['diagonally-button'][-1], 0,0,2,2)

    main_grid.addWidget(main_widget['ost-logo-label'][-1], 3,0,3,3)
    pixmap = QPixmap('img/OST_LOGO_.png')
    scaled_pixmap = pixmap.scaled(QSize(40,40), Qt.KeepAspectRatio, Qt.SmoothTransformation)
    main_widget['ost-logo-label'][-1].setPixmap(scaled_pixmap)

    main_widget['ERP-code-button'][-1].setText('')
    main_grid.addWidget(main_widget['ERP-code-button'][-1], 6,0,4,3)
    
    main_widget['search-button'][-1].setText('')
    main_grid.addWidget(main_widget['search-button'][-1], 10,0,4,3)

    main_widget['stock-button'][-1].setText('')
    main_grid.addWidget(main_widget['stock-button'][-1], 14,0,4,3)

    main_widget['MRP-button'][-1].setText('')
    main_grid.addWidget(main_widget['MRP-button'][-1], 18,0,4,3)
    
    main_grid.removeItem(main_widget['ERP-code-layout'][-1])
    main_grid.addLayout(main_widget['ERP-code-layout'][-1], 5, 3, 4, 7)
    
def toggle_sub_layout(sub_layout):
    # 切換子項目按鈕的顯示和隱藏狀態
    if sub_layout.itemAt(0).widget().isVisible():
        hide_sub_layout(sub_layout)
    else:
        show_sub_layout(sub_layout)

def show_sub_layout(sub_layout):
    # 顯示子項目按鈕
    for i in range(sub_layout.count()):
        widget = sub_layout.itemAt(i).widget()
        widget.show()

def hide_sub_layout(sub_layout):
    # 隱藏子項目按鈕
    for i in range(sub_layout.count()):
        widget = sub_layout.itemAt(i).widget()
        widget.hide()