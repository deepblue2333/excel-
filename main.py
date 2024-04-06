import csv
import sys

from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QTableWidgetItem, QApplication, QButtonGroup, QTextBrowser

from MainWindow_ui import Ui_MainWindow
from ChangeExcelMessage import change_cake_first


class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MyMainWindow, self).__init__()
        self.setupUi(self)
        self.row_num = 20  # 默认生成20行
        self.report_data = None

        self.blank_template = [['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
                               ['', '', '', '', '', '', '', '', '', '', '', '', '', '']]

        self.cake_first_template = [
            ['产品名称', '批量', '抽 (送) 样日期', '商标', '型号规格', '生产日期 (批号)', '样品数量',
             '验讫日期', '水分，％', '细菌总数 (CFU/g)', '大肠菌群 (CFU/g)', '时间', '审核人', '主检人'],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', '']]

        # 创建一个QButtonGroup并将RadioButton添加到其中
        self.button_group = QButtonGroup(self)
        self.button_group.addButton(self.radioButton)
        self.button_group.addButton(self.radioButton_2)
        self.button_group.addButton(self.radioButton_3)
        self.button_group.addButton(self.radioButton_4)

        self.pushButton_2.clicked.connect(self.close_window)

        self.pushButton.clicked.connect(self.update_dataframe)

        # 连接QButtonGroup的buttonClicked信号到槽函数
        self.button_group.buttonClicked.connect(self.on_radio_button_clicked)

    def on_radio_button_clicked(self, button):
        # 槽函数，处理RadioButton的点击事件
        print(f'RadioButton "{button.text()}" clicked')

        self.show_dataframe(self.blank_template)
        if button.text() == "蛋糕第一份":
            self.show_dataframe(self.cake_first_template)

    def generate_report(self):
        # 槽函数，处理Submit按钮的点击事件
        selected_button = self.button_group.checkedButton()
        if selected_button:
            print(f'Submit clicked. Selected RadioButton: {selected_button.text()}')
        else:
            print('Submitted clicked. No RadioButton selected.')

    def show_message_box(self):
        # 模拟程序执行完成后弹出消息框
        QMessageBox.information(self.parent(), "任务完成", "程序执行完成！", QMessageBox.Ok)

    def close_window(self):
        self.close()

    def show_dataframe(self, data):
        # 设置表格的行列数
        self.tableWidget.setRowCount(self.row_num)
        self.tableWidget.setColumnCount(len(data[0]))

        # 设置表头
        self.tableWidget.setHorizontalHeaderLabels(data[0])

        # 填充数据
        for row in range(self.row_num):
            for col in range(len(data[0])):
                item = QTableWidgetItem(str(data[row + 1][col]))
                self.tableWidget.setItem(row, col, item)

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            selected_text = ""

            for item in self.tableWidget.selectedItems():
                selected_text += "{}\n".format(item.text())

            selected_text = selected_text.rstrip('\t')

            clipboard = QApplication.clipboard()
            clipboard.setText(selected_text)

        elif event.matches(QKeySequence.Paste):
            clipboard = QApplication.clipboard()
            paste_text = clipboard.text()

            # 直接处理粘贴文本，将 \r\n（CRLF）和 \n（LF）作为分隔符
            rows = [row.split('\r\n') if '\r\n' in row else row.split('\n') for row in paste_text.split('\t')]

            for i, value in enumerate(rows):
                if "=" in value[0]:
                    if self.tableWidget.currentRow() + i < self.tableWidget.rowCount() and \
                            self.tableWidget.currentColumn() < self.tableWidget.columnCount():
                        text_browser = QTextBrowser()
                        text_browser.setPlainText('\n'.join(value).strip("\""))
                        self.tableWidget.setCellWidget(self.tableWidget.currentRow(),
                                                       self.tableWidget.currentColumn() + i,
                                                       text_browser)
                else:
                    if self.tableWidget.currentRow() + i < self.tableWidget.rowCount() and \
                            self.tableWidget.currentColumn() < self.tableWidget.columnCount():
                        item = QTableWidgetItem(value[0])
                        self.tableWidget.setItem(self.tableWidget.currentRow(),
                                                 self.tableWidget.currentColumn() + i,
                                                 item)

    def update_dataframe(self):

        # 创建一个空列表来存储所有行的数据
        table_data = []

        # 获取tableWidget的行数和列数
        num_rows = self.tableWidget.rowCount()
        num_cols = self.tableWidget.columnCount()

        # 遍历每一行
        for row in range(num_rows):
            # 创建一个空列表来存储当前行的数据
            row_data = []

            # 遍历当前行的每一列
            for col in range(num_cols):
                # 获取当前单元格的CellWidget
                cell_widget = self.tableWidget.cellWidget(row, col)

                if cell_widget is not None:
                    # 如果存在CellWidget，进一步处理以获取文本值
                    if isinstance(cell_widget, QTextBrowser):
                        # 如果是 QTextBrowser，获取纯文本值
                        cell_value = cell_widget.toPlainText()
                    else:
                        # 其他类型的 CellWidget 的处理方式
                        # 根据具体情况调整
                        cell_value = str(cell_widget.text())  # 示例中假设是一个 widget 具有 text() 方法
                else:
                    # 如果单元格没有 CellWidget，继续使用 QTableWidgetItem 的逻辑
                    item = self.tableWidget.item(row, col)
                    if item is not None:
                        cell_value = item.text()
                    else:
                        cell_value = ""  # 如果单元格为空，则将其值设置为空字符串

                # 将单元格的值添加到当前行的数据列表中
                row_data.append(cell_value)

            # 将当前行的数据列表添加到主列表中
            table_data.append(row_data)

        self.report_data = table_data

        print("Rename Files:")
        print(table_data)

        for row in table_data:
            if row[0] != '':
                change_cake_first(row)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MyMainWindow()
    main_window.show()
    sys.exit(app.exec_())
