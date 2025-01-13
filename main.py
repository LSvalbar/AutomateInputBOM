import sys
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QLineEdit, QFileDialog, \
    QMessageBox, QHBoxLayout
from DrissionPage import Chromium, ChromiumOptions


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
        self.path_input.setGeometry(0, 0, 200, 15)
        self.path_input.setPlaceholderText("请输入Excel文件路径")
        sub1_layout.addWidget(self.path_input, alignment=Qt.AlignVCenter)

        # 创建浏览按钮来选择文件
        self.browse_button = QPushButton("浏览文件", self)
        self.browse_button.clicked.connect(self.browse_file)
        sub1_layout.addWidget(self.browse_button, alignment=Qt.AlignVCenter)

        total_layout.addLayout(sub1_layout)
        # 确认按钮
        self.confirm_button = QPushButton("确认", self)
        self.confirm_button.clicked.connect(self.on_confirm)
        sub2_layout.addWidget(self.confirm_button, alignment=Qt.AlignCenter)

        # 关闭按钮
        self.close_button = QPushButton("关闭", self)
        self.close_button.clicked.connect(self.on_close)
        sub2_layout.addWidget(self.close_button, alignment=Qt.AlignCenter)

        total_layout.addLayout(sub2_layout)
        self.setLayout(total_layout)

    # Step 1: 读取 Excel 文件
    def read_excel(self, file_path):
        total_list = []
        df = pd.read_excel(file_path)
        product_num_list = df.iloc[:, 0].tolist()
        graphy_num_list = df.iloc[:, 1].tolist()
        version_num_list = df.iloc[:, 2].tolist()
        total_list.append(product_num_list)
        total_list.append(graphy_num_list)
        total_list.append(version_num_list)
        return total_list

    # Step 2: 使用 DrissionPage 执行自动化任务
    def automate_browser(self, data):
        co = ChromiumOptions()
        co.set_argument('--start-maximized')
        browser_tab = Chromium(co).latest_tab
        try:
            # 打开目标页面
            browser_tab.get("http://192.168.0.21:8080/xbiot_fsd_mes")
            browser_tab.wait.doc_loaded()
            # 账户名
            account_ele = browser_tab.ele('x://html/body/div/div[2]/div[2]/form/div[1]/input')
            account_ele.input('D4116')
            browser_tab.wait(1,1.5)
            # 密码
            password_ele = browser_tab.ele('x://html/body/div/div[2]/div[2]/form/div[2]/input')
            password_ele.input('123')
            browser_tab.wait(1, 1.5)
            # 登录
            login_btn_ele = browser_tab.ele('x://html/body/div/div[2]/div[2]/form/div[3]/a')
            login_btn_ele.click()
            browser_tab.wait.doc_loaded()
            # 基础信息
            foundation_info_ele = browser_tab.ele('x://html/body/div[2]/div[2]/div[4]/div[1]')
            #browser_tab.wait.eles_loaded('x://html/body/div[2]/div[2]/div[4]/div[1]')
            browser_tab.wait.doc_loaded()
            foundation_info_ele.wait.clickable()
            foundation_info_ele.click()
            # 产品信息
            product_info_ele = browser_tab.ele('x://html/body/div[2]/div[2]/div[4]/div[2]/div[3]')
            #browser_tab.wait.eles_loaded('x://html/body/div[2]/div[2]/div[4]/div[2]/div[3]')
            browser_tab.wait.doc_loaded()
            product_info_ele.wait.clickable()
            product_info_ele.click()
            browser_tab.wait.doc_loaded()
            #browser_tab.wait.eles_loaded('x://html/body/div[1]/div[2]/button[1]')
            # 新增
            new_add_ele = browser_tab.ele('x://html/body/div[1]/div[2]/button[1]')
            for count in range(len(data[0])):
                # 新增
                new_add_ele.wait.clickable()
                new_add_ele.click()
                browser_tab.wait.eles_loaded('x://html/body/div[4]/div/div[2]/div/div[1]/div/input')
                # 产品编号
                product_num_ele = browser_tab.ele('x://html/body/div[4]/div/div[2]/div/div[1]/div/input')
                #product_num_ele.wait.enabled()
                product_num_ele.input(data[0][count])
                # 图号
                graphy_num_ele = browser_tab.ele('x://html/body/div[4]/div/div[2]/div/div[2]/div/input')
                #graphy_num_ele.wait.enabled()
                graphy_num_ele.click(data[1][count])
                # 产品名称
                product_name_ele = browser_tab.ele('x://html/body/div[4]/div/div[2]/div/div[3]/div/input')
                #product_name_ele.wait.enabled()
                product_name_ele.input(data[0][count])
                # 重量上限
                weight_max_ele = browser_tab.ele('x://html/body/div[4]/div/div[2]/div/div[4]/div/input')
                #weight_max_ele.wait.enabled()
                weight_max_ele.input(0)
                # 重量下限
                weight_min_ele = browser_tab.ele('x://html/body/div[4]/div/div[2]/div/div[5]/div/input')
                #weight_min_ele.wait.enabled()
                weight_min_ele.input(0)
                # 版本号
                version_no_ele = browser_tab.ele('x://html/body/div[4]/div/div[2]/div/div[6]/div/input')
                #version_no_ele.wait.enabled()
                version_no_ele.input(data[2][count])
                # 保存
                save_btn_ele = browser_tab.ele('x://html/body/div[4]/div/div[2]/div/div[7]/button')
                save_btn_ele.wait.enabled()
                save_btn_ele.click()
                browser_tab.wait.load_start()
        except Exception as error:
            self.show_error_message(error.__str__())
        finally:
            pass

    def browse_file(self):
        # 弹出文件选择对话框，选择 Excel 文件
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
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
