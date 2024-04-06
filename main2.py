import csv
import sys

from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QTableWidgetItem, QApplication, QButtonGroup, QTextBrowser

from MainWindow2 import Ui_MainWindow


class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MyMainWindow, self).__init__()
        self.setupUi(self)
        self.row_num = 20  # 默认生成20行
        self.report_data = None

        # 创建一个QButtonGroup并将RadioButton添加到其中
        self.button_group = QButtonGroup(self)
        self.button_group.addButton(self.radioButton)
        self.button_group.addButton(self.radioButton_2)
        self.button_group.addButton(self.radioButton_3)
        self.button_group.addButton(self.radioButton_4)

        self.pushButton_2.clicked.connect(self.close_window)

        # 连接QButtonGroup的buttonClicked信号到槽函数
        self.button_group.buttonClicked.connect(self.on_radio_button_clicked)

    def on_radio_button_clicked(self, button):
        # 槽函数，处理RadioButton的点击事件
        print(f'RadioButton "{button.text()}" clicked')

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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MyMainWindow()
    main_window.show()
    sys.exit(app.exec_())
