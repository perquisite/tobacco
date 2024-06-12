# my_prefix = r"D:\烟草项目\tobacco-xiejunyu\tobacco-xiejunyu\yancaoRegularDemo\副本"
# prefix = r'D:\烟草项目\tobacco-8.7\yancaoRegularDemo\副本'
import win32com
from win32com.client import Dispatch, constants
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.ReadFile import DocxData
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
from yancaoRegularDemo.Resource.tools.TimeOperator import TimeOper
import os
import time

from yancaoRegularDemo.Resource.tools.get_pictures import *
from yancaoRegularDemo.Resource.tools.utils import *
from yancaoRegularDemo.Resource.tools.OCR_IDCard import *

class Table_60(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prefix = source_prefix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        # self.doc = self.mw.Documents.Open(self.my_prefix + "含身份证文书_.docx")
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.check_NameLoc_IDcard()",
            "self.check_Validity_IDcard()"
        ]

    # 检查 身份证姓名和住址 与《行政处罚决定书》中的当事人和住址保持一致
    def check_NameLoc_IDcard(self):
        file_name = '行政处罚决定书'
        if os.path.exists(self.source_prefix + file_name + "_.docx") == 0:
            table_father.display(self, "文件缺失：《" + file_name + "》不存在", "red")
        else:
            # 先提取《行政处罚决定书》中的当事人姓名
            other_tabels_content = DocxData(self.source_prefix + file_name + "_.docx").text
            self_name_parttern = re.compile(r'[(当事人)]：\s*([^\s]*)\s*')
            self_address_parttern = re.compile(r'[(住址)]：\s*([^\s]*)\s*')
            self_name = re.findall(self_name_parttern, other_tabels_content)
            self_address = re.findall(self_address_parttern, other_tabels_content)
            other_name = self_name[0]
            other_address = self_address[0]
            print(other_name)
            print(other_address)
            # 再提取含身份证图片文书中的图片，进行识别提取
            #get_pictures_single(self.source_prefix)
            pic_path_1 = self.source_prefix + r"picture\含身份证文书_\word\media\image1.png"
            print(pic_path_1)
            id_ocr = OCR_IDCard(pic_path_1, "")
            id_name = id_ocr.getName()
            id_address = id_ocr.getAddress()
            print(id_name)
            print(id_address)
            # 判断姓名、地址是否一致
            if id_name == other_name:
                pass
            else:
                table_father.display(self, "从身份证图片提取的姓名与“当事人”一栏不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '', '从身份证图片提取出姓名与“当事人”一栏不一致')

            if id_address == other_address:
                pass
            else:
                table_father.display(self, "从身份证图片提取的地址与“地址”一栏不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '', '从身份证图片提取的地址与“地址”一栏不一致')

    # 检查 身份证到期时间应在提取时间之后
    def check_Validity_IDcard(self):
        pic_path_2 = self.source_prefix + r"picture\含身份证文书_\word\media\image2.png"
        print(pic_path_2)
        id_ocr = OCR_IDCard("", pic_path_2)
        id_date = id_ocr.getExpiringDate()
        print(id_date)
        #print(type(id_date))
        # 转化日期表达格式
        str_list = list(id_date)
        str_list.insert(4, '-')
        str_list.insert(7, '-')
        id_date = ''.join(str_list)
        t = TimeOper()
        if t.time_order(id_date, t.getLocalDate()) >= 0:
            pass
        else:
            table_father.display(self, "身份证已过期！", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '', '身份证已过期！')

    # def check(self, contract_file_path):
    #     print("正在审查《含身份证文书》，审查结果如下：")
    #     data = DocxData(file_path = contract_file_path)
    #     self.contract_text = data.text
    #     self.contract_tables_content = data.tabels_content
    #     for func in self.all_to_check:
    #         #try:
    #         eval(func)
    #         #except Exception as e:
    #         #table_father.display(self, "文档存在格式错误，函数失效：" + func + ' 遇到错误:' + str(e.args))
    #     self.doc.Close()
    #     # self.mw.Quit()
    #     print("《含身份证文书》审查完毕\n")
    #     info_list_result = table_father.get_info_list(self)
    #     return info_list_result

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
        self.mw = win32com.client.Dispatch("Word.Application")
        self.doc = self.mw.Documents.Open(self.my_prefix + file_name_real)
        data = DocxData(file_path=contract_file_path)
        self.contract_text = data.text
        self.contract_tables_content = data.tabels_content
        for func in self.all_to_check:
            try:
                eval(func)
            except Exception as e:
                table_father.display(self, "文档存在格式错误，函数失效：" + func + ' 遇到错误:' + str(e.args))
        self.doc.Close()
        # self.mw.Quit()
        print(file_name_real + "审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = r"C:\Users\twj\Desktop\test" + '\\'
    #print(my_prefix)
    list1 = os.listdir(my_prefix)
    #if "行政处罚决定书_.docx" in list1:
    ioc = Table_60(my_prefix, my_prefix)
    contract_file_path = my_prefix + r"含身份证文书_.docx"
    #print(contract_file_path)
    ioc.check(contract_file_path, r"含身份证文书_.docx")



