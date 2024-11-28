import sys,time,os
from PyQt5.QtWidgets import (QApplication, QWidget,QDesktopWidget, QVBoxLayout, QLineEdit, QFileDialog, QStatusBar,
                            QGridLayout, QPushButton, QMessageBox,QProgressBar, QFrame, QLabel, QHBoxLayout, QTextEdit,
                            QTableWidget, QTableWidgetItem, QGroupBox,QHeaderView)
from PyQt5.QtCore import Qt, QTimer,pyqtSignal
from PyQt5.QtGui import QFont,QIcon
import subprocess
import json


class SuccessDialog(QMessageBox):
    def __init__(self, file_path):
        super().__init__()
        self.setWindowTitle("生成成功")
        self.resize(800, 500)  # 设置窗口大小
        # 将相对路径转换为绝对路径
        self.file_path = os.path.abspath(file_path)
        
        # 设置显示内容为绝对路径
        self.setText(f"文件路径: {self.file_path}")

        # 添加按钮
        self.addButton("打开文件", QMessageBox.ActionRole)
        self.addButton("打开文件所在位置", QMessageBox.ActionRole)
        self.addButton("关闭", QMessageBox.RejectRole)

        # 连接按钮的点击事件
        self.buttonClicked.connect(self.handle_button_click)
        self.setFont(QFont('Arial', 14))

    def handle_button_click(self, button):
        if button.text() == "打开文件":
            self.open_file()
        elif button.text() == "打开文件所在位置":
            self.open_file_location()
        elif button.text() == "关闭":
            self.close()

    def open_file(self):
        # 尝试打开文件
        if os.path.exists(self.file_path):
            os.startfile(self.file_path)
            # self.close()  # 打开文件后关闭对话框
        else:
            QMessageBox.warning(self, "错误", "文件不存在!", QMessageBox.Ok)

    def open_file_location(self):
        # 打开文件所在文件夹并选中文件
        file_folder = os.path.dirname(self.file_path)
        if os.path.exists(file_folder):
            subprocess.run(['explorer', '/select,', self.file_path])
        else:
            QMessageBox.warning(self, "错误", "文件夹不存在!", QMessageBox.Ok)



class Select_data_Window(QWidget):
    data_sent = pyqtSignal(int)
    def __init__(self,data):
        super().__init__()
        self.setWindowTitle("选择需要生成的小组")
        self.resize(1500, 800)  # 设置窗口大小
        # 创建一个垂直布局来放置表格和按钮
        layout = QVBoxLayout()

        # 创建一个 QTableWidget
        self.table =QTableWidget(0,2)
        self.table.setHorizontalHeaderLabels(["小组负责人", "作业项目"])  # 设置表头
        #填充数据
        for row, item in enumerate(data):
            self.table.insertRow(row)
            self.table.setRowHeight(row, 120)
            for col, value in enumerate(item):
                self.table.setItem(row, col, QTableWidgetItem(str(value)))
        self.table.setFont(QFont('Arial', 14))
        # 使表格支持选择行
        self.table.setSelectionBehavior(QTableWidget.SelectRows)

        # 设置表格列宽自适应，平分剩余空间
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)  # 第一列占满一半
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # 第二列占满另一半

        # 创建确定按钮
        self.confirm_button = QPushButton("确定", self)
        self.confirm_button.clicked.connect(self.confirm_selection)
        self.confirm_button.setFont(QFont('Arial', 14))

        # 将表格和按钮添加到布局中
        layout.addWidget(self.table)
        layout.addWidget(self.confirm_button)

        # 设置窗口的布局
        self.setLayout(layout)

    def confirm_selection(self):
        # 获取选中的行
        self.select_index = -1
        selected_rows = self.table.selectedIndexes()
        if not selected_rows:
            QMessageBox.warning(self, "警告", "请先选择一行数据！", QMessageBox.Ok)
            return

        # 取第一个选中的行索引
        self.select_index = selected_rows[0].row()
        self.data_sent.emit(self.select_index)
        self.close()

class initWindow(QWidget):
    data_sent = pyqtSignal(int)
    def __init__(self):
        super().__init__()
        self.resize(500, 100)  # 设置窗口大小为 300x100 px
        self.setWindowFlags(Qt.FramelessWindowHint)  # 去除窗口的标题栏和边框
        self.setStyleSheet("background-color: #f0f0f0;")  # 设置背景色为淡灰色

        self.label = QLabel("初始化中：读取数据模板...", self)
        self.label.setFont(QFont('Arial', 14))
        self.label.setAlignment(Qt.AlignCenter)  # 文本居中
        self.label.resize(500, 100)  # 设置文本框的大小，和窗口一致

        screen_geometry = QApplication.desktop().availableGeometry()
        window_x = (screen_geometry.width() - self.width()) // 2
        window_y = (screen_geometry.height() - self.height()) // 2
        self.move(window_x, window_y)  # 将窗口放置在屏幕中央
        QTimer.singleShot(500, self.run_init)  # 调用 run_init
        self.show()

    def run_init(self):
        try:
            self.writer = main_z.init_("./派工单数据模板.xlsx") #初始化模板
            self.enter_second_interface()
        except:
            QMessageBox.warning(self, "警告", "数据模板读取失败，请检查数据模板是否存在", QMessageBox.Ok)
            QApplication.quit() #直接退出
        

    def enter_second_interface(self):
        self.second_window = mainWindow(self.writer)
        self.second_window.show()
        self.close()  # 关闭当前窗口


class FileDropArea(QGroupBox):
    def __init__(self, title):
        super().__init__(title)
        self.setAcceptDrops(True)
        self.setStyleSheet("border: 2px dashed lightgray; padding: 10px;")
        self.setFont(QFont('Arial', 12))
        
        self.label = QLabel("拖入文件或选择文件...")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setFont(QFont('Arial', 14))
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        self.setLayout(layout)

        # 用于存储文件路径
        self.file_path = ""
        self.data = ()
        self.file_type = title

    def alert(self,msg):
        msg_box = QMessageBox()
        msg_box.setWindowTitle('警告')
        msg_box.setText(msg)
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        file_path = [url.toLocalFile() for url in event.mimeData().urls()][-1]
        if os.path.isfile(file_path):
            self.add_file(file_path)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel表格 (*.xls *.xlsx);;所有文件(*)")
            if os.path.isfile(file_path):
                self.add_file(file_path)


    def check_data(self):

        #检查并读入数据
        try:
            if self.file_path and self.file_type == "维修方案":
                self.data = main_z.read_fix_input(self.file_path) #读入维修方案
            if self.file_path and self.file_type == "工队班次":
                self.data = main_z.read_worker_input(self.file_path) #读入工队班次
        except Exception as e:
            print(e)

            self.alert("读取失败，请检查:%s"%self.file_type)
            self.label.setText("拖入文件或选择文件...")
            self.file_path = ""


    def add_file(self, path):
        #文本框ui中修改显示
        self.file_path = path  # 添加文件路径到列表
        file_name = os.path.basename(path)  # 只获取文件名
        self.label.setText(file_name)
        self.check_data()

    def get_data(self):
        return self.data

    def get_file_path(self):
        if not self.file_path:
            self.alert("请先选择文件:%s"%self.file_type)
            return ""
        else:
            return self.file_path  # 返回文件路径



class data_source(QGroupBox):
    def __init__(self,):
        super().__init__("数据来源")
        # self.setAcceptDrops(True)
        # self.setStyleSheet("border: 2px dashed lightgray; padding: 10px;")
        self.setFont(QFont('Arial', 12))
        
        self.excel_button = QPushButton("来自Excel")#, self)
        self.json_button = QPushButton("来自Json")#, self)
        self.excel_button.setMinimumHeight(50)  # 设置按钮最小高度为50像素
        self.excel_button.setMaximumHeight(200)  # 设置最大高度
        self.json_button.setMinimumHeight(50)  # 设置按钮最小高度为50像素
        self.json_button.setMaximumHeight(200)  # 设置最大高度
        self.excel_button.clicked.connect(self.select_excel)
        self.json_button.clicked.connect(self.select_json)
        layout = QHBoxLayout()
        layout.addWidget(self.excel_button)
        layout.addWidget(self.json_button)
        self.setLayout(layout)

        self.path = ""
        self.data = None

    def select_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel表格 (*.xls *.xlsx);;所有文件(*)")
        if os.path.isfile(file_path):
            self.path = file_path
        else:
            return

    def select_json(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Json文件", "", "Json表格 (*.json);;所有文件(*)")
        if os.path.isfile(file_path):
            self.path = file_path
        else:
            return
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
            # 使用 json.load() 方法加载文件中的 JSON 数据
                self.data = json.load(file)
            QMessageBox.information(self, "提示", "读取成功", QMessageBox.Ok)
            return self.data
        except Exception as e:
            QMessageBox.warning(self, "警告", "读取失败。\n"+str(e), QMessageBox.Ok)


class template_path(QGroupBox):
    def __init__(self,):
        super().__init__("模板")
        # self.setAcceptDrops(True)
        # self.setStyleSheet("border: 2px dashed lightgray; padding: 10px;")
        self.setFont(QFont('Arial', 12))
        
        self.template = QPushButton("选择模板文件")
        self.template.setMinimumHeight(50)  # 设置按钮最小高度为50像素
        self.template.setMaximumHeight(200)  # 设置最大高度
        self.template.clicked.connect(self.template_select)
        layout = QHBoxLayout()
        layout.addWidget(self.template)
        self.setLayout(layout)
        self.path = ""
    def template_select(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择模板文件", "", "Excel/Word (*.xls *.xlsx *.doc *.docx);;所有文件(*)")
        if os.path.isfile(file_path):
            self.path = file_path


class mainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.writer = ""
        self.setWindowTitle("DocForge")
        self.setGeometry(100, 100, 800, 400)
        layout = QHBoxLayout()
        self.data_source_area = data_source()
        self.template_path_area = template_path()

        left_layout = QVBoxLayout()
        left_layout.addWidget(self.data_source_area)
        left_layout.addWidget(self.template_path_area)

        layout.addLayout(left_layout)

        # 开始按钮
        self.start_button = QPushButton("开始")
        # self.start_button.setMinimumHeight(50)  # 设置按钮最小高度为50像素
        self.start_button.clicked.connect(self.select)
        self.start_button.setFont(QFont('Arial', 14))

        self.status_bar = QStatusBar(self)
        self.status_bar.showMessage('作者：https://github.com/kakasearch')
        self.status_bar.setFixedHeight(50)  # 设置按钮高度为50


        # 整体布局
        overall_layout = QVBoxLayout()
        overall_layout.addLayout(layout)
        overall_layout.addWidget(self.start_button)
        overall_layout.addWidget(self.status_bar)

        screen_geometry = QApplication.desktop().availableGeometry()
        window_x = (screen_geometry.width() - self.width()) // 2
        window_y = (screen_geometry.height() - self.height()) // 2
        self.move(window_x, window_y)  # 将窗口放置在屏幕中央
        self.setLayout(overall_layout)

    def handle_main(self,row):
        fix_add_data = main_z.re_get(r"维修方案(.*)\.xls",self.fix_data_area.get_file_path())
        result_path = main_z.main(self.fix_data,row,self.workers,self.tools,self.assigns,self.writer,fix_add_data)
        dialog = SuccessDialog(result_path)
        dialog.exec_()


    def select(self):
        if( not self.worker_data_area.get_file_path()) or (not self.fix_data_area.get_file_path()):
            return
        self.fix_data = self.fix_data_area.get_data()
        self.workers,self.tools,self.assigns = self.worker_data_area.get_data()
        select_data = [ [x["小组负责人及手持号"] ,x["作业项目"].strip()] for x in self.fix_data]
        self.Select_data_Window = Select_data_Window(select_data)
        self.Select_data_Window.data_sent.connect(self.handle_main)
        self.Select_data_Window.show()
        

data = {
    "a":[{"n1":1},{"n2":23}],
    "b":123
}
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('icon.png'))
    window = mainWindow()
    window.show()
    sys.exit(app.exec_())
















