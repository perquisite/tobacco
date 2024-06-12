# -*- coding:utf-8 -*-
# @Author: tanweijia
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

from yancaoRegularDemo.Resource.tools.tanweijia_function import is_exist_cover
from yancaoRegularDemo.Resource.tools.utils import *
import win32com
from win32com.client import Dispatch, constants
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh

function_description_dict = {
    'check_All': '1、送达文书名称：行政处罚决定书'
                 '2、送达文书文号：与《行政处罚决定书》编号一致 ' 
                 '3、受送达人：与《行政处罚决定书》中的当事人一致'
                 '4、送达地点：不为空'
                 '5、送达方式：应当为直接送达、留置送达、邮寄送达、委托送达、公告送达中的其中一种。',
}

# 送达回证
class Table_39(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prefix = source_prefix # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        # self.doc = self.mw.Documents.Open(self.my_prefix + "送达回证_.docx")
        self.contract_text = None
        self.contract_tables_content = None
        

        self.all_to_check = [
            "self.check_All()"
        ]

    # 检查 送达文书名称和送达文书文号
    def check_All(self):
        text0 = self.contract_tables_content["送达文书名称"]
        text1 = self.contract_tables_content["送达文书文号"]
        text2 = self.contract_tables_content["受送达人"]
        text3 = self.contract_tables_content["送达地点"]
        text4 = self.contract_tables_content["送达方式"]
        #text5 = self.contract_tables_content["收件人签名或盖章"]
        #text6 = self.contract_tables_content["送达人签名"]

        if "行政处罚决定书" not in text0:
            table_father.display(self, "送达文书名称：送达文书名称错误，应为 行政处罚决定书", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "名    称", "送达文书名称错误，应为 行政处罚决定书")
        final_name = is_exist_cover("行政处罚决定书", "当场行政处罚决定书", self.source_prefix)
        if not final_name:
            table_father.display(self, "文书缺失：" + "《行政处罚决定书》.docx不存在", "red")
        else:
            data = DocxData(self.source_prefix + final_name)
            temp = data.text.split()
            num = temp[3].strip()
            print(num)
            if num != text1:
                table_father.display(self, "送达文书文号：送达文书文号与《行政处罚决定书》编号（"+num+"）不一致！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "文    号", "送达文书文号与《行政处罚决定书》编号（"+num+"）不一致")
            # temp0 = re.search("[决定书编号：]\s*([^\s]*)\s*", data.text)
            # if temp0:
            #     temp1 = temp0.group(0).strip()
            #     print(temp1)
            #     print(text1)
            #     if temp1 != text1:
            #         table_father.display(self, "× 送达文书文号与《行政处罚决定书》编号不一致！", "red")
            #         tyh.addRemarkInDoc(self.mw, self.doc, "文    号", "送达文书文号与《行政处罚决定书》编号不一致")
            #     reference_text = data.text
            # else:
            #     table_father.display(self, "× 没有在《行政处罚决定书》中找到其编号！", "red")
            text_temp = re.findall(".*当事人：(.*?)\n", data.text)
            #print(text_temp[0])
            if temp:
                if text_temp[0] != text2:
                    table_father.display(self, "受送达人：受送达人与《行政处罚决定书》中的当事人（"+text_temp[0]+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "受送达人", "受送达人与《行政处罚决定书》中的当事人（"+text_temp[0]+"）不一致!")
            else:
                table_father.display(self, "受送达人：没有提取到《行政处罚决定书》中的当事人，因此无法与“受送达人”比较请人工核查！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "受送达人", "受送达人：没有提取到《行政处罚决定书》中的当事人，因此无法与“受送达人”比较请人工核查！")

        if text3 == "":
            table_father.display(self, "送达地点：”送达地点“为空！", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "送达地点", "”送达地点“不应为空！")

        if text4 not in ["直接送达", "留置送达", "邮寄送达", "委托送达", "公告送达"]:
            table_father.display(self, "送达方式：送达方式”不规范！应当为直接送达、留置送达、邮寄送达、委托送达、公告送达中的其中一种。", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "送达方式", "”送达方式“不规范，应当为直接送达、留置送达、邮寄送达、委托送达、公告送达中的其中一种。")

        # if not text5:           # ？不做
        #     table_father.display(self,"× 收件人签名为空！", "red")
        #     tyh.addRemarkInDoc(self.mw, self.doc, "收件人签名", "收件人签名 不应为空！")

        # text6 送达人签名          签名图片需求不做
        # if not text6:
        #     table_father.display(self,"× 送达人签名为空！", "red")
        #     tyh.addRemarkInDoc(self.mw, self.doc, "送达人签名", "送达人签名 不应为空！")
        # else:
        #     if "," in text6:
        #         name_list = text6.split(",")
        #     elif "，" in text6:
        #         name_list = text6.split("，")
        #     elif "、" in text6:
        #         name_list = text6.split("、")
        #     else:
        #         name_list = text6.split()   #或许可以调用命名实体识别
        #     print(name_list)
        #     if len(name_list) < 2:
        #         table_father.display(self, "× 送达人签名 应当为两人以上！", "red")
        #         tyh.addRemarkInDoc(self.mw, self.doc, "送达人签名", "送达人签名 应当为两人以上")

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
                table_father.display(self,
                                     "文档格式有误，请主观审查下列功能：" + function_description_dict[str(func)[5:-2]],
                                     "red")
                table_father.display(self, "文档存在格式错误，函数失效：" + func + ' 遇到错误:' + str(e.args))
        self.doc.Close()
        self.mw.Quit()
        print(file_name_real + "审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\twj\Desktop\\test\\"
    list = os.listdir(my_prefix)
    if "送达回证_.docx" in list:
        ioc = Table_39(my_prefix, my_prefix)
        contract_file_path = my_prefix + "送达回证_.docx"
        ioc.check(contract_file_path, "送达回证_.docx")