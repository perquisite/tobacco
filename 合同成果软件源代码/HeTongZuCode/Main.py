# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""

from PyQt5.QtCore import pyqtSlot, QBasicTimer, QTimerEvent
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox
from PyQt5.QtWidgets import QFileDialog
from ContractType import ContractType

from UiLayout import Ui_MainWindow

from DocReader import DocReader

from over_all_description import get_over_all_file

over_all_info = []


class MainWindow(QMainWindow, Ui_MainWindow):
    """
    Class documentation goes here.
    """
    docs_to_process = None
    need_to_export = True
    export_dir = None
    status = "ready"  # 当前程序是否是就绪状态（无需要进行的任务）

    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget (defaults to None)
        @type QWidget (optional)
        """
        super(MainWindow, self).__init__(parent)
        self.timer = QBasicTimer()
        self.setupUi(self)
        self.step = 0
        self.step_size = None

    def timerEvent(self, a0: 'QTimerEvent') -> None:

        global over_all_info

        if self.step >= self.step_size:
            print("全部审查完毕。。。")
            self.status = "ready"
            self.step_size = -1  # 目的是避免重复点击“审查文件”按钮
            self.pushButton_2.setText('开始审查文件')
            # self.textBrowser.append("全部审查完毕。。。")
            QMessageBox.information(self, u"提示信息", u"全部审查完毕！")
            self.timer.stop()
            self.step = 0

            get_over_all_file(over_all_info, self.export_dir)
            return

        self.status = "work"
        file = self.docs_to_process[self.step]
        print("-----------------------------------\n正在审查文件：" + file)
        self.textBrowser.append("-----------------------------------\n正在审查文件：" + file)
        d = DocReader(file, self.export_dir)
        contract_check_result = d.to_info()
        if contract_check_result == "docx_blank":
            print("请确定\"" + file + "\"不是空文档！")
            self.textBrowser.append("请确定\"" + file + "\"不是空文档！")
            over_all_info.append("空白" + file)

        # if contract_check_result.type == "type_not_sure":
        if contract_check_result.type == ContractType.NotSure:
            print("合同类型未匹配成功！请检查文档内容是否合乎规范！！")
            # self.textBrowser.append("合同类型未匹配成功！请检查文档内容是否合乎规范！！")
            self.textBrowser.append("合同为非标准合同，对该合同进行风险性检测和完整性检测！")
            over_all_info.append([file,contract_check_result])

        if contract_check_result != "docx_blank" and contract_check_result != "type_not_sure" and contract_check_result.type != ContractType.NotSure:
            print("错误的要素:\n" + str(contract_check_result.factors_error))
            print("提示的要素:\n" + str(contract_check_result.factors_to_inform))
            self.textBrowser.append("错误的要素:\n" + str(contract_check_result.factors_error))
            self.textBrowser.append("提示的要素:\n" + str(contract_check_result.factors_to_inform))
            over_all_info.append([file, contract_check_result])
        self.step = self.step + 1
        self.progressBar.setValue(int(self.step / self.step_size * 100))

    # "选择待审查文件" 按钮按下
    @pyqtSlot()
    def on_pushButton_clicked(self):
        """
        Slot documentation goes here.
        """
        my_file_path, _ = QFileDialog.getOpenFileNames(self, u"打开文件", '/', "word(*.docx *.doc)")
        if len(my_file_path) == 0:
            self.textBrowser_2.setText("您取消了操作！")
            # self.tex
            return
        self.textBrowser_2.setText("您选择了以下文件：")
        for dir in my_file_path:
            self.textBrowser_2.append(dir)
        self.docs_to_process = my_file_path
        self.step_size = len(my_file_path)

    # "开始审查文件" 按钮按下
    @pyqtSlot()
    def on_pushButton_2_clicked(self):
        """
        Slot documentation goes here.
        """
        if self.status == "ready":
            self.textBrowser.clear()

            if self.docs_to_process == None:
                QMessageBox.information(self, u"提示信息", u"请添加要审查的文件！")
                return

            if self.need_to_export and self.export_dir == None:
                QMessageBox.information(self, u"提示信息", u"请设置审查结果导出路径，或者取消勾选“导出审查结果文件”。\n因为您勾选了导出审查文件，但未设置路径！")
                return

            if self.step_size == -1:
                # print("任务已经完成！无需重复点击审查。\n若需要审查新的文件,请重新选择待审查的新文件")
                # self.textBrowser.append("任务已经完成！无需重复点击审查。\n若需要审查新的文件,请重新选择待审查的新文件")
                QMessageBox.information(self, u"提示信息", u"任务已经完成！无需重复点击审查。\n若需要审查新的文件,请重新选择待审查的新文件")
                return

        if self.timer.isActive():
            self.textBrowser.append("#####################已暂停！")
            self.timer.stop()
            self.pushButton_2.setText('继续')
        else:
            self.textBrowser.append("已开始！")
            self.timer.start(100, self)
            self.pushButton_2.setText('暂停')

        # self.textBrowser.clear()
        # for file in self.docs_to_process:
        #     print("-----------------------------------\n正在审查文件：" + file)
        #     self.textBrowser.append("-----------------------------------\n正在审查文件：" + file)
        #     d = DocReader(file)
        #     contract_check_result = d.to_info()
        #
        #     if contract_check_result == "docx_blank":
        #         print("请确定\"" + file + "\"不是空文档！")
        #         self.textBrowser.append("请确定\"" + file + "\"不是空文档！")
        #
        #     if contract_check_result == "type_not_sure":
        #         print("合同类型未匹配成功！请检查文档内容是否合乎规范！！")
        #         self.textBrowser.append("合同类型未匹配成功！请检查文档内容是否合乎规范！！")
        #
        #     if contract_check_result != "docx_blank" and contract_check_result != "type_not_sure":
        #         print("错误的要素:\n" + str(contract_check_result.factors_error))
        #         print("提示的要素:\n" + str(contract_check_result.factors_to_inform))
        #         self.textBrowser.append("错误的要素:\n" + str(contract_check_result.factors_error))
        #         self.textBrowser.append("提示的要素:\n" + str(contract_check_result.factors_to_inform))
        #
        #     d.remove_new_file()
        # print("全部审查完毕。。。")
        # self.textBrowser.append("全部审查完毕。。。")

    # "导出审查结果文件" CheckBox按下
    @pyqtSlot()
    def on_checkBox_clicked(self):
        """
        Slot documentation goes here.
        """
        print(self.checkBox.isChecked())
        self.need_to_export = self.checkBox.isChecked()

    # "设置导出路径" 按钮按下
    @pyqtSlot()
    def on_pushButton_3_clicked(self):
        """
        Slot documentation goes here.
        """
        get_directory_path = QFileDialog.getExistingDirectory(self,
                                                              "选取指定文件夹",
                                                              "/")
        self.lineEdit.setText(get_directory_path)
        self.export_dir = get_directory_path


if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    ui = MainWindow()
    ui.show()
    sys.exit(app.exec_())
