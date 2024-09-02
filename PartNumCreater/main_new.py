import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QScrollArea, QListWidget, QListWidgetItem, QSplitter
from PyQt5.QtCore import Qt

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 設置主窗口的佈局
        main_layout = QHBoxLayout(self)

        # 創建導航欄
        nav_bar = QListWidget()
        nav_bar.addItem(QListWidgetItem("Home"))
        nav_bar.addItem(QListWidgetItem("Page 1"))
        nav_bar.addItem(QListWidgetItem("Page 2"))
        nav_bar.addItem(QListWidgetItem("Page 3"))
        nav_bar.setFixedWidth(200)  # 設置導航欄寬度
        nav_bar.setStyleSheet("background-color: lightgray;")  # 設置導航欄背景顏色

        # 創建可滾動區域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        # 創建滾動區域內部的widget
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)

        # 添加幾個有背景顏色的標籤以方便確認
        for i in range(20):
            label = QLabel(f"Label {i+1}")
            label.setStyleSheet("background-color: lightblue; border: 1px solid black;")
            label.setFixedHeight(50)  # 固定高度
            scroll_layout.addWidget(label)

        scroll_area.setWidget(scroll_content)

        # 使用QSplitter來調整布局
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(nav_bar)
        splitter.addWidget(scroll_area)
        splitter.setSizes([200, 600])  # 設置導航欄和滾動區域的初始比例

        # 將splitter添加到主佈局
        main_layout.addWidget(splitter)

        self.setLayout(main_layout)
        self.setWindowTitle('PyQt Navigation and Scroll Area Example')
        self.resize(800, 600)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec_())