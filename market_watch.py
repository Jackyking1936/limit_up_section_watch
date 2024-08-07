from login_gui import LoginForm

import sys
import json
import math
from fubon_neo.sdk import FubonSDK, Mode
import pandas as pd
from pathlib import Path
import pickle
from datetime import datetime

from PySide6.QtWidgets import QTabWidget, QTableWidgetItem, QFileDialog, QApplication, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QGridLayout, QLabel, QLineEdit, QPushButton, QSizePolicy, QPlainTextEdit
from PySide6.QtGui import QIcon, QTextCursor, QColor
from PySide6.QtCore import Qt, Signal, QObject, QSize


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

    def setText(self, text):
        super().setText(text)
        self.value = self._convert_to_number(text)

class Communicate(QObject):
    # 定義一個帶參數的信號
    print_log_signal = Signal(str)
    table_update_signal = Signal(str, str, int, str)
    color_update_signal = Signal(str, str, bool)

class MainApp(QWidget):
    def __init__(self, active_account):
        super().__init__()

        self.active_account = active_account

        my_icon = QIcon()
        my_icon.addFile('market.png') 
        self.setWindowIcon(my_icon)
        self.setWindowTitle("Python漲停濾網看盤(教學範例，使用前請先了解相關內容)")
        self.resize(1200, 700)

        # 製作上下排列layout上為庫存表，下為log資訊
        layout = QVBoxLayout()

        # 庫存表表頭
        self.info_header = ['股票名稱', '股票代號', '市場別', '開盤價','最高價','最低價', '現價', '漲幅(%)', '9:40前漲停']
        
        self.info_tab = QTabWidget()
        table = QTableWidget(0, len(self.info_header))
        table.setHorizontalHeaderLabels(self.info_header)
        self.info_tab.addTab(table, '類股名稱')

        # 整個設定區layout
        layout_parameter = QGridLayout()

        # 參數設置區layout設定

        label_file_path = QLabel('目標清單路徑:')
        label_file_path.setStyleSheet("QLabel { font-weight: bold; }")
        layout_parameter.addWidget(label_file_path, 0, 1, Qt.AlignRight)

        self.lineEdit_default_file_path = QLineEdit()
        layout_parameter.addWidget(self.lineEdit_default_file_path, 0, 2, 1, 4)
        self.folder_btn = QPushButton('')
        self.folder_btn.setIcon(QIcon('folder.png'))
        layout_parameter.addWidget(self.folder_btn, 0, 6)
        self.folder_btn.setFixedSize(QSize(32, 32))

        self.read_excel_btn = QPushButton('')
        self.read_excel_btn.setText('讀取清單')
        self.read_excel_btn.setStyleSheet("QPushButton { font-weight: bold; }")
        layout_parameter.addWidget(self.read_excel_btn, 0, 7)

        for i in range(8, 11):
            label_dummy = QLabel('')
            label_dummy.setStyleSheet("QLabel { font-weight: bold; }")
            layout_parameter.addWidget(label_dummy, 0, i)
        
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)

        layout.addWidget(self.info_tab)
        layout.addLayout(layout_parameter)
        layout.addWidget(self.log_text)

        layout.setStretchFactor(self.info_tab, 7)
        layout.setStretchFactor(layout_parameter, 1)
        layout.setStretchFactor(self.log_text, 3)

        self.setLayout(layout)

        my_target_path = None
        my_target_list_file = Path("./target_list_path.pkl")
        if my_target_list_file.is_file():
            with open('./target_list_path.pkl', 'rb') as f:
                temp_dict = pickle.load(f)
                my_target_path = temp_dict['target_list_path']

        # Open the file dialog to select a file
        if my_target_path:
            self.lineEdit_default_file_path.setText(my_target_path)

        self.print_log("login success, 現在使用帳號: {}".format(self.active_account.account))
        self.print_log("建立行情連線...")

        sdk.init_realtime(Mode.Normal) # 建立買進行情連線
        self.reststock = sdk.marketdata.rest_client.stock
        self.websocket = sdk.marketdata.websocket_client.stock

        # button slot function
        self.folder_btn.clicked.connect(self.showDialog)
        self.read_excel_btn.clicked.connect(self.read_watch_list)

        # communicator slot function
        self.communicator = Communicate()
        self.communicator.print_log_signal.connect(self.print_log)
        self.communicator.table_update_signal.connect(self.update_table)
        self.communicator.color_update_signal.connect(self.limit_up_coloring)

        # default parameter initilaize
        self.subscribed_ids = {}
        self.col_idx_map = dict(zip(self.info_header, range(len(self.info_header))))
        threshold_time = datetime.today().replace(hour=9, minute=40, second=0, microsecond=0)
        self.threshold_unix = int(datetime.timestamp(threshold_time)*1000000)

        # websocket connect
        self.websocket.on("message", self.handle_message)
        self.websocket.on("connect", self.handle_connect)
        self.websocket.on("disconnect", self.handle_disconnect)
        self.websocket.on("error", self.handle_error)
        self.websocket.connect()

    def limit_up_coloring(self, table_name, symbol, is_limit_up):
        if is_limit_up:
            self.table_name_maps[table_name].item(self.row_symbol_maps[table_name][symbol], self.col_idx_map['漲幅(%)']).setBackground(QColor(Qt.red))
            self.table_name_maps[table_name].item(self.row_symbol_maps[table_name][symbol], self.col_idx_map['漲幅(%)']).setForeground(QColor(Qt.white))
        else:
            self.table_name_maps[table_name].item(self.row_symbol_maps[table_name][symbol], self.col_idx_map['漲幅(%)']).setBackground(QColor(Qt.transparent))
            self.table_name_maps[table_name].item(self.row_symbol_maps[table_name][symbol], self.col_idx_map['漲幅(%)']).setForeground(QColor())

    def update_table(self, table_name, symbol, col, value):
        try:
            table = self.table_name_maps[table_name]
            table.item(self.row_symbol_maps[table_name][symbol], col).setText(value)
            if col == self.col_idx_map['漲幅(%)']:
                table.sortByColumn(col, Qt.DescendingOrder)
                symbol_list = []
                for i in range(table.rowCount()):
                    symbol_list.append(table.item(i, self.col_idx_map['股票代號']).text())
                self.row_symbol_maps[table_name] = dict(zip(symbol_list, range(len(symbol_list))))
        except Exception as e:
            print(e, self.row_symbol_maps[table_name][symbol], col, value)


    def handle_message(self, message):
        msg = json.loads(message)
        event = msg["event"]
        data = msg["data"]
        # print(event, data)
        
        # subscribed事件處理
        if event == "subscribed":
            if type(data) == list:
                # print(event, data)
                for sub_record in data:
                    self.communicator.print_log_signal.emit('訂閱成功...'+sub_record['symbol'])
                    self.subscribed_ids[sub_record['symbol']] = sub_record['id']
            else:
                id = data["id"]
                symbol = data["symbol"]
                self.communicator.print_log_signal.emit('訂閱成功...'+symbol)
                self.subscribed_ids[symbol] = id
        
        # subscribed事件處理
        elif event == "unsubscribed":
            if type(data) == list:
                print(event, data)
            else:
                id = data["id"]
                symbol = data["symbol"]
                self.communicator.print_log_signal.emit('訂閱成功...'+symbol)
                self.subscribed_ids[symbol] = id

        # subscribed事件處理
        elif event == "snapshot":
            # print(event, data)
            symbol = data['symbol']
            market_type = data['market']
            open_price = data['openPrice']
            high_price = data['highPrice']
            low_price = data['lowPrice']
            cur_price = data['lastPrice']
            change_percent = data['changePercent']

            for name in self.table_name_maps.keys():
                # print(name)
                if symbol in self.row_symbol_maps[name]:
                    if 'isLimitUpPrice' in data:
                        if data['isLimitUpPrice']:
                            self.communicator.color_update_signal.emit(name, symbol, data['isLimitUpPrice'])
                    self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['市場別'], str(market_type))
                    self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['開盤價'], str(open_price))
                    self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['最高價'], str(high_price))
                    self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['最低價'], str(low_price))
                    self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['現價'], str(cur_price))
                    self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['漲幅(%)'], str(change_percent)+'%')

        elif event == "data":
            # print(event, data)
            if 'isTrial' in data:
                if data['isTrial']:
                    return
                
            symbol = data['symbol']
            high_price = data['highPrice']
            low_price = data['lowPrice']
            cur_price = data['lastPrice']
            change_percent = data['changePercent']
            tick_time = data['lastUpdated']

            for name in self.table_name_maps.keys():
                # print(name)
                if symbol in self.row_symbol_maps[name]:
                    if 'isLimitUpPrice' in data:
                        if data['isLimitUpPrice']:
                            if tick_time<self.threshold_unix:
                                self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['9:40前漲停'], 'Y')
                            self.communicator.color_update_signal.emit(name, symbol, data['isLimitUpPrice'])
                        else:
                            self.communicator.color_update_signal.emit(name, symbol, False)

                    self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['最高價'], str(high_price))
                    self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['最低價'], str(low_price))
                    self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['現價'], str(cur_price))
                    self.communicator.table_update_signal.emit(name, symbol, self.col_idx_map['漲幅(%)'], str(change_percent)+'%')

    def handle_connect(self):
        self.communicator.print_log_signal.emit('market data connected')
    
    def handle_disconnect(self, code, message):
        self.communicator.print_log_signal.emit(f'market data disconnect: {code}, {message}')
    
    def handle_error(self, error):
        self.communicator.print_log_signal.emit(f'market data error: {error}')

    def read_watch_list(self):
        file_path = self.lineEdit_default_file_path.text()
        watch_df = pd.read_excel(file_path)
        self.column_names = [col_name for col_name in watch_df.columns if 'Unnamed' not in col_name]
        self.table_dict = {}
        self.table_name_maps = {}
        self.row_symbol_maps = {}

        for i, col_name in enumerate(self.column_names):
            self.table_dict[col_name] = watch_df.iloc[:, (2*i):(2*i+2)]
            self.table_dict[col_name].columns = ['代碼', '名稱']
            self.table_dict[col_name] = self.table_dict[col_name].iloc[1:, :]
            self.table_dict[col_name] = self.table_dict[col_name].dropna(axis=0, how = 'all')
            self.table_dict[col_name] = self.table_dict[col_name].reset_index(drop=True)

            rows = self.table_dict[col_name].shape[0]
            cols = len(self.info_header)
            table = QTableWidget(rows, cols)
            table.setHorizontalHeaderLabels(self.info_header)

            self.table_name_maps[col_name] = table
            self.row_symbol_maps[col_name] = {}

            # ['股票名稱', '股票代號', '市場別', '開盤價','最高價','最低價', '現價', '漲幅(%)', '9:40前漲停']
            for i in range(rows):
                for j in range(cols):
                    item = QTableWidgetItem('-')
                    if self.info_header[j] == '股票名稱':
                        item.setText(self.table_dict[col_name].iloc[i, 1])
                    elif self.info_header[j] == '股票代號':
                        symbol = self.table_dict[col_name].iloc[i, 0].replace('.TW', '')
                        item.setText(symbol)
                        self.row_symbol_maps[col_name][symbol] = i
                    elif self.info_header[j] == '漲幅(%)':
                        item = NumericTableWidgetItem('-')
                        

                    table.setItem(i, j, item)

            self.info_tab.addTab(table, col_name)
            self.websocket.subscribe({
                'channel': 'aggregates', 
                'symbols': list(self.row_symbol_maps[col_name].keys())
            })
        self.info_tab.removeTab(0)

    def showDialog(self):
        my_target_path = None
        my_target_list_file = Path("./target_list_path.pkl")
        if my_target_list_file.is_file():
            with open('./target_list_path.pkl', 'rb') as f:
                temp_dict = pickle.load(f)
                my_target_path = temp_dict['target_list_path']

        # Open the file dialog to select a file
        if my_target_path:
            file_path, _ = QFileDialog.getOpenFileName(self, '請選擇您的觀察目標清單', my_target_path, 'All Files (*)')
        else:
            file_path, _ = QFileDialog.getOpenFileName(self, '請選擇您的觀察目標清單', 'C:\\', 'All Files (*)')

        if file_path:
            self.lineEdit_default_file_path.setText(file_path)
            temp_dict = {
                'target_list_path': file_path
            }
            with open('target_list_path.pkl', 'wb') as f:
                pickle.dump(temp_dict, f)

    # 更新最新log到QPlainTextEdit的slot function
    def print_log(self, log_info):
        self.log_text.appendPlainText(log_info)
        self.log_text.moveCursor(QTextCursor.End)

    # 視窗關閉時要做的事，主要是關websocket連結及存檔現在持有部位
    def closeEvent(self, event):
        
        self.print_log("disconnect websocket...")
        self.websocket.disconnect()
        sdk.logout()

        can_exit = True
        if can_exit:
            event.accept() # let the window close
        else:
            event.ignore()

if __name__ == "__main__":
    try:
        sdk = FubonSDK()
    except ValueError:
        raise ValueError("請確認網路連線")
    
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()
    app.setStyleSheet("QWidget{font-size: 12pt;}")
    form = LoginForm(MainApp, sdk, 'market.png')
    form.show()
    
    sys.exit(app.exec())