import os
import sys
import shutil
import time

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from MainWindow import *
from zhishitupu.tools.simple_content import *
from zhishitupu.tools.transform import *
from function import *

template_src_path = getRootPath() + "模板.xlsx"


class MyWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)

        self.pushButton.clicked.connect(self.addButton)  # “添加”按钮
        self.pushButton_2.clicked.connect(self.deleteButton)  # “删除”按钮
        self.pushButton_3.clicked.connect(self.editButton)  # “修改”按钮

    @pyqtSlot()
    def editButton(self):
        text1 = self.lineEdit.text()
        text2 = self.lineEdit_2.text()
        s = Simple_Content()
        if not s.is_null(text1) and not s.is_null(text2):
            rdf_path = getRootPath() + "rdf"
            city_path = rdf_path + "\\" + text1
            city_rdf_list = get_allfile(rdf_path)
            if city_path not in city_rdf_list:  # 没有这个城市的存档（文件夹）
                QMessageBox.information(self, u"提示信息", u"库中无此城市的相关信息。您可以先尝试添加。")
            else:  # 有这个城市的存档（文件夹）的话
                rdf_list = get_allfile(city_path)
                should_edit_rdf_path = city_path + "\\" + text2 + ".rdf"
                if should_edit_rdf_path not in rdf_list:  # 该城市文件夹下没有同名rdf，无法修改，提示可以添加
                    QMessageBox.information(self, u"提示信息", u"库中无此城市的该法条。您可以先尝试添加。")
                else:  # 有了同名rdf，直接修改。
                    # 先打开缓存
                    cache_path = getRootPath() + "缓存文件夹" + "\\" + text1 + "第" + text2 + "条.xlsx"
                    os.startfile(cache_path)
                    result = QMessageBox.information(self, u"提示信息", u"请编辑打开的excel文件，保存并关闭后请点击”OK“")
                    while True:
                        if result == QtWidgets.QMessageBox.Ok:
                            excle_to_rdf(cache_path, should_edit_rdf_path)
                            QMessageBox.information(self, u"提示信息", u"已将信息加入库中")
                            self.lineEdit.setText('')
                            self.lineEdit_2.setText('')
                            break

    def addButton(self):
        text1 = self.lineEdit.text()
        text2 = self.lineEdit_2.text()
        s = Simple_Content()
        if not s.is_null(text1) and not s.is_null(text2):
            rdf_path = getRootPath() + "rdf"  # rdf库的路径
            city_path = rdf_path + "\\" + text1  # rdf库中的城市文件夹路径（不管事先存不存在该城市文件夹）
            city_rdf_list = get_allfile(rdf_path)
            # create_folder = getRootPath() + "\\" + "rdf" + "\\" + text1  # 应创建的城市对应文件夹
            create_rdf_path = city_path + "\\" + text2 + ".rdf"  # 应创建的rdf路径
            template_dst_path = getRootPath() + r"\缓存文件夹" + "\\" + text1 + "第" + text2 + "条" + ".xlsx"  # 缓存文件夹的路径名
            if city_path not in city_rdf_list:  # 没有这个城市的存档（文件夹），那么直接添加（先创建文件夹在创建rdf）
                copy_and_start(template_src_path, template_dst_path)
                result = QMessageBox.information(self, u"提示信息", u"请编辑打开的excel文件，保存并关闭后请点击”OK“")
                while True:
                    if result == QtWidgets.QMessageBox.Ok:
                        os.mkdir(city_path)  # 创建对应城市文件夹
                        excle_to_rdf(template_dst_path, create_rdf_path)
                        #os.remove(template_dst_path)  # 清除缓存文件 （万一不关闭，需要报错处理）
                        QMessageBox.information(self, u"提示信息", u"已将信息加入库中")
                        self.lineEdit.setText('')
                        self.lineEdit_2.setText('')
                        break
            else:  # 有这个城市的存档（文件夹）的话
                rdf_list = get_allfile(city_path)
                should_add_rdf_path = city_path + "\\" + text2 + ".rdf"
                if should_add_rdf_path not in rdf_list:  # 该城市文件夹下没有同名rdf，那么添加（直接添加rdf）
                    copy_and_start(template_src_path, template_dst_path)
                    result = QMessageBox.information(self, u"提示信息", u"请编辑打开的excel文件，保存并关闭后请点击”OK“")
                    while True:
                        if result == QtWidgets.QMessageBox.Ok:
                            excle_to_rdf(template_dst_path, create_rdf_path)
                            # os.remove(template_dst_path)  # 清除缓存文件 （万一不关闭，需要报错处理）
                            QMessageBox.information(self, u"提示信息", u"已将信息加入库中")
                            self.lineEdit.setText('')
                            self.lineEdit_2.setText('')
                            break
                else:  # 有了同名rdf，则不能添加，提示可以选择修改。
                    QMessageBox.information(self, u"提示信息", u"该城市的库中已存在此法条。您可以尝试修改。")

    def deleteButton(self):  # 已只存在城市文件夹，但还要删除的时候？
        text1 = self.lineEdit.text()
        text2 = self.lineEdit_2.text()
        s = Simple_Content()
        if not s.is_null(text1) and not s.is_null(text2):
            rdf_path = getRootPath() + "rdf"
            city_path = rdf_path + "\\" + text1
            city_rdf_list = get_allfile(rdf_path)
            if city_path in city_rdf_list:  # 有这个城市的存档（文件夹）
                rdf_list = get_allfile(city_path)
                should_delete_rdf_path = city_path + "\\" + text2 + ".rdf"
                cache_path = getRootPath() + "缓存文件夹" + "\\" + text1 + "第" + text2 + "条.xlsx"
                if should_delete_rdf_path in rdf_list:
                    os.remove(should_delete_rdf_path)
                    os.remove(cache_path)
                    QMessageBox.information(self, u"提示信息", u"删除完毕！")
                    self.lineEdit.setText('')
                    self.lineEdit_2.setText('')
                else:
                    QMessageBox.information(self, u"提示信息", u"该城市的库中不存在此法条。您可以尝试先添加。")
            else:
                QMessageBox.information(self, u"提示信息", u"库中不含该城市信息。您可以尝试先添加。")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWin = MyWindow()
    myWin.show()
    sys.exit(app.exec_())
