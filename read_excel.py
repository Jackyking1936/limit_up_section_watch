#%%
import pandas as pd

watch_df = pd.read_excel('類股清單.xlsx')
column_names = [col_name for col_name in watch_df.columns if 'Unnamed' not in col_name]
table_dict = {}

for i, col_name in enumerate(column_names):
    table_dict[col_name] = watch_df.iloc[:, (2*i):(2*i+2)]
    table_dict[col_name].columns = ['代碼', '名稱']
    table_dict[col_name] = table_dict[col_name].iloc[1:, :]
    table_dict[col_name] = table_dict[col_name].dropna(axis=0, how = 'all')
    table_dict[col_name] = table_dict[col_name].reset_index(drop=True)
# %%
from PySide6.QtWidgets import QTableWidgetItem
from PySide6.QtCore import Qt

class NumericTableWidgetItem(QTableWidgetItem):
    def __init__(self, value):
        super().__init__(str(value))
        self.value = self._convert_to_number(value)

    def _convert_to_number(self, value):
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            # 處理百分比
            if value.strip() in ['-', '', 'N/A']:
                return float('-inf')  # 使用負無窮大表示這些特殊值
            
            if value.endswith('%'):
                try:
                    return float(value.rstrip('%')) / 100
                except ValueError:
                    pass
            # 處理普通數字
            try:
                return float(value.replace(',', ''))  # 處理千位分隔符
            except ValueError:
                pass
        return value

    def __lt__(self, other):
        if isinstance(other, NumericTableWidgetItem):
            if isinstance(self.value, (int, float)) and isinstance(other.value, (int, float)):
                return self.value < other.value
        return super().__lt__(other)
    
def populate_table(table_widget, data):
    table_widget.setRowCount(len(data))
    table_widget.setColumnCount(len(data[0]))
    
    for row, row_data in enumerate(data):
        for col, value in enumerate(row_data):
            table_widget.setItem(row, col, NumericTableWidgetItem(value))

# 示例用法
from PySide6.QtWidgets import QApplication, QTableWidget
import sys

app = QApplication(sys.argv)

table_widget = QTableWidget()
data = [
    [1000, "50%", -3.14],
    [-2000, "25", "-6%"],
    [7000.5, "75%", "9%"],
    [500, "-33.33%", '9%']
]
populate_table(table_widget, data)
table_widget.setSortingEnabled(True)

# 設置列標題
table_widget.setHorizontalHeaderLabels(["整數", "百分比", "混合"])

# 調整列寬以適應內容
table_widget.resizeColumnsToContents()
table_widget.sortByColumn(1, Qt.DescendingOrder)

table_widget.show()
sys.exit(app.exec())  