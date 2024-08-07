from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                               QTableWidget, QTableWidgetItem, QTabWidget, QHeaderView)
from PySide6.QtCore import Qt

class ChromeStyleTabWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        main_layout = QVBoxLayout(self)

        # 创建标签页小部件
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # 创建并添加不同的表格页面
        self.addTableTab("表格 1", 5, 3, ["列 A", "列 B", "列 C"])
        self.addTableTab("表格 2", 4, 2, ["名称", "值"])
        self.addTableTab("表格 3", 6, 4, ["ID", "名称", "数量", "价格"])

        self.setLayout(main_layout)
        self.setWindowTitle('Chrome 风格标签页表格')
        self.setGeometry(100, 100, 800, 600)

    def addTableTab(self, name, rows, cols, headers):
        table = self.createTable(rows, cols, headers)
        self.tab_widget.addTab(table, name)

    def createTable(self, rows, cols, headers):
        table = QTableWidget(rows, cols)
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # 填充一些示例数据
        for i in range(rows):
            for j in range(cols):
                table.setItem(i, j, QTableWidgetItem(f"数据 {i+1},{j+1}"))

        return table

if __name__ == '__main__':
    app = QApplication([])
    ex = ChromeStyleTabWidget()
    ex.show()
    app.exec()
