import sys
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QLineEdit, QFileDialog, \
    QMessageBox, QHBoxLayout
from DrissionPage import Chromium,ChromiumOptions

# Step 3: 创建 PyQt 界面
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("输入 Excel 文件路径")
        self.setGeometry(200, 200, 500, 300)
        self.setAutoFillBackground(True)

        total_layout = QVBoxLayout()
        sub1_layout = QHBoxLayout()
        sub2_layout = QHBoxLayout()

        # 创建文本框用于输入文件路径
        self.path_input = QLineEdit(self)
        self.path_input.setGeometry(0,0,200,15)
        self.path_input.setPlaceholderText("请输入Excel文件路径：")
        sub1_layout.addWidget(self.path_input,alignment=Qt.AlignVCenter)

        # 创建浏览按钮来选择文件
        self.browse_button = QPushButton("浏览文件", self)
        self.browse_button.clicked.connect(self.browse_file)
        sub1_layout.addWidget(self.browse_button,alignment=Qt.AlignVCenter)

        total_layout.addLayout(sub1_layout)
        # 确认按钮
        self.confirm_button = QPushButton("确认", self)
        self.confirm_button.clicked.connect(self.on_confirm)
        sub2_layout.addWidget(self.confirm_button,alignment=Qt.AlignCenter )

        # 关闭按钮
        self.close_button = QPushButton("关闭", self)
        self.close_button.clicked.connect(self.on_close)
        sub2_layout.addWidget(self.close_button,alignment=Qt.AlignCenter )

        total_layout.addLayout(sub2_layout)
        self.setLayout(total_layout)

    # Step 1: 读取 Excel 文件
    def read_excel(self,file_path):
        total_list = []
        df = pd.read_excel(file_path)
        page_num_index = df.iloc[:, 0].tolist()
        product_num_index = df.iloc[:, 1].tolist()
        product_name_index = df.iloc[:, 2].tolist()
        total_list.append(page_num_index)
        total_list.append(product_num_index)
        total_list.append(product_name_index)
        return total_list

    # Step 2: 使用 DrissionPage 执行自动化任务
    def automate_browser(self,data):
        co = ChromiumOptions()
        co.set_argument('--start-maximized')
        browser_tab = Chromium(co).latest_tab
        str = ['狂飙','人民的名义','三国演义','红楼梦','西游记']
        try:
            # 打开目标页面
            browser_tab.get("http://www.baidu.com")
            browser_tab.wait.load_start()
            for count in range(len(data[0])):
                #input_ele = browser_tab.ele('x://html/body/div[1]/div[1]/div[5]/div/div/form/span[1]/input')
                input_ele = browser_tab.ele('#kw')
                input_ele.wait.enabled()
                input_ele.clear()
                input_ele.wait(0.5, 1)
                input_ele.input(str[count])
                input_ele.wait(0.5,1)
                search_btn = browser_tab.ele('@value=百度一下')
                search_btn.wait.clickable()
                search_btn.click()
        except Exception as error:
            self.show_error_message(error)
        finally:
            pass

    def browse_file(self):
        # 弹出文件选择对话框，选择 Excel 文件
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.path_input.setText(file_path)  # 将选中的文件路径显示在文本框中

    def on_confirm(self):
        # 获取输入的文件路径
        file_path = self.path_input.text()
        if not file_path:
            self.show_error_message("文件路径不能为空！")
            return

        # 读取 Excel 文件中的数据
        data = self.read_excel(file_path)

        # 执行自动化任务
        self.close()  # 关闭当前窗口
        self.automate_browser(data)


        # 程序执行完成后弹出完成窗口
        self.show_complete_message()

    def on_close(self):
        # 关闭程序
        self.close()

    def show_complete_message(self):
        # 程序执行完成后弹出提示框
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("执行完成")
        msg.setText("程序执行已完成！")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def show_error_message(self, message):
        # 弹出错误提示框
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("错误")
        msg.setText(message)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

# 主程序入口
def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
