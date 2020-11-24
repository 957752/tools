from PySide2.QtGui import QIcon
from PySide2.QtWidgets import QApplication, QMessageBox, QFileDialog
from PySide2.QtUiTools import QUiLoader
import xlrd


class Stats:

    def __init__(self):
        # 从文件中加载UI定义
        self.ui = QUiLoader().load('url.ui')
        # 选择导入表格
        self.ui.pushButton_4.clicked.connect(self.sum)
        # 清除功能
        self.ui.pushButton_3.clicked.connect(self.clear)
        # 获取表格
        self.ui.pushButton.clicked.connect(self.get_file)


    def clear(self):
        self.ui.plainTextEdit.clear()

    def get_file(self):
        filePath, _ = QFileDialog.getOpenFileName(
            self.ui,  # 父窗口对象
            "请选择Excel表格",  # 标题
            r"C:\Users\Administrator\Desktop",  # 起始目录
            "文件类型 (*.xlsx *.xls *.xml)"  # 选择类型过滤项，过滤内容在括号中
        )

        self.ui.lineEdit.setText(filePath)

    def str_of_num(self, num):
        if num < 1000000:
            return '{}{}'.format(0, "亿")
        else:
            result = num / 100000000
            return '{}{}'.format(round(result, 3), "亿")

    def sum(self):
        path = self.ui.lineEdit.text()
        if path == "":
            QMessageBox.critical(
                self.ui,
                '错误',
                '请先选择选择Excel表格！')
        else:
            book = xlrd.open_workbook(path)

            # 得到所有sheet对象

            sheet = book.sheet_by_index(0)

            # 全面合同库，57行，只保留泰行审xxx
            licence = sheet.col_values(colx=56, start_rowx=1)

            # 积存数据
            dataAccumulation = 0
            for lic, lic_1 in enumerate(licence):
                if lic_1.startswith('泰行审') and (sheet.cell_value(lic, 35) == '未售' or sheet.cell_value(lic, 35) == '保留'):
                    income = sheet.cell_value(lic, 63)
                    dataAccumulation += income

            # 洋房数据
            dataVilla = 0
            for lic, lic_1 in enumerate(licence):
                if lic_1.startswith('泰行审') and (sheet.cell_value(lic, 35) == '未售' or sheet.cell_value(lic, 35) == '保留') and sheet.cell_value(lic, 8) == '洋房':
                    income = sheet.cell_value(lic, 63)
                    dataVilla += income

            # 商铺数据
            dataShops = 0
            for lic, lic_1 in enumerate(licence):
                if lic_1.startswith('泰行审') and (sheet.cell_value(lic, 35) == '未售' or sheet.cell_value(lic, 35) == '保留') and sheet.cell_value(lic, 8) == '商铺':
                    income = sheet.cell_value(lic, 63)
                    dataShops += income

            dataShops = self.str_of_num(dataShops)
            dataVilla = self.str_of_num(dataVilla)
            dataAccumulation = self.str_of_num(dataAccumulation)

            self.ui.plainTextEdit.appendPlainText(f"截止目前项目积存{dataAccumulation}，其中：洋房{dataVilla}、商铺{dataShops}；")


app = QApplication([])


app.setWindowIcon(QIcon('12.jpg'))
stats = Stats()
stats.ui.show()
app.exec_()

