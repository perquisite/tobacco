# -*- coding:utf-8 -*-
# @Author: tanweijia
import win32com
from win32com.client import Dispatch, constants
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.ReadFile import DocxData
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
# from yancaoRegularDemo.twj_ReadFile import File_39
from yancaoRegularDemo.Resource.tools.utils import *

# 不予行政处罚决定书
class Table_38(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prefix = source_prefix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.contract_text = None
        self.contract_tables_content = None

    def check(self, contract_file_path, file_name_real):
        contract_file_path = self.source_prefix
        #print("正在审查《不予行政处罚决定书》，审查结果如下：")
        if tyh.file_exists(contract_file_path, "不予行政处罚决定书"):
            #print("存在《不予行政处罚决定书》!")
            table_father.display(self, "有无《不予行政处罚决定书》：" + "本案卷中存在《不予行政处罚决定书》!", "red")
        else:
            pass
            #print("不存在《不予行政处罚决定书》!")
        #print("《不予行政处罚决定书》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result

if __name__ == '__main__':
    source_prefix = "C:\\Users\\twj\Desktop\\test\\"
    print(source_prefix)
    ioc = Table_38(source_prefix, source_prefix)
    contract_file_path = source_prefix #+ "不予行政处罚决定书_.docx"
    # print(contract_file_path)
    ioc.check(contract_file_path, "不予行政处罚决定书_.docx")