# -*- coding:utf-8 -*-
# @Author: tanweijia
import win32com
from win32com.client import Dispatch, constants
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.ReadFile import DocxData
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os

from yancaoRegularDemo.Resource.tools.EntityRecognition import EntityRecognition
from yancaoRegularDemo.Resource.tools.tanweijia_function import is_exist_cover
from yancaoRegularDemo.Resource.tools.utils import *

function_description_dict = {
    'check_Name_Time_and_Loc': '1、抬头应与《行政处罚决定书》中的当事人一致；2、时间应当与《立案报告表》中记载的“案发时间”一致；3、地点应当与《立案报告表》中记载的“案发地点”一致。',
    'check_reason2': '依据应当与《案件处理审批表》中的“处罚依据”一致。',
    'check_reason1': '案由应当与《立案报告表》中记载的“案由”一致。',
    'check_Opinion_and_Content': '1.决定应当与《案件处理审批表》中的“承办人意见”一致; 2.送达的内容应当为“行政处罚决定书”+编号。',
}

# my_prefix = r"D:\烟草项目\tobacco-xiejunyu\tobacco-xiejunyu\yancaoRegularDemo\副本"
# prefix = r'D:\烟草项目\tobacco-8.7\yancaoRegularDemo\副本'
# 送达公告
class Table_40(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prefix = source_prefix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.check_Name_Time_and_Loc()",
            "self.check_reason2()",
            "self.check_reason1()",
            "self.check_Opinion_and_Content()"
        ]

    # 检查 抬头、时间、地点
    def check_Name_Time_and_Loc(self):
        # Name
        #data1 = DocxData(self.source_prefix + "送达公告_.docx")
        final_name = is_exist_cover("行政处罚决定书", "当场行政处罚决定书", self.source_prefix)
        if not final_name:
            table_father.display(self, "文书缺失：" + "《行政处罚决定书》.docx不存在", "red")
        else:
            data = DocxData(self.source_prefix + final_name)
            reference_text = data.text
            # print(reference_text)
            text_temp = re.findall(".*当事人：(.*?)\n", reference_text)
            # print(text_temp)
            temp0 = self.contract_text.split("\n")
            text1 = temp0[3].strip()
            text1 = text1.replace("：", "")
            # print(text1)
            if text_temp:
                if not text_temp[0] in text1:
                    table_father.display(self, "抬头：”抬头“与《行政处罚决定书》中的当事人（"+text_temp[0]+"）不一致!", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, text1, "”抬头“与《行政处罚决定书》中的当事人（"+text_temp[0]+"）不一致!")
            else:
                table_father.display(self, "未从《行政处罚决定书》中提取到“当事人”，因此无法与“抬头”比较，请人工核查！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, text1, "未从《行政处罚决定书》中提取到“当事人”，因此无法与“抬头”比较，请人工核查！")

        # Time and Loc
        final_name = is_exist_cover("立案报告表", "撤销立案报告表", self.source_prefix)
        if not final_name:
            pass
            # table_father.display(self, "文书缺失：" + "《立案报告表》.docx不存在", "red")
        else:
            data = DocxData(self.source_prefix + final_name)
            # Time
            text_temp = self.contract_text.split()[-1]
            #print(text_temp)
            a = tyh.changeDate(text_temp)
            #print(a)
            b = tyh.get_strtime(data.tabels_content["案发时间"])  # 立案报告表中的时间
            #print(b)
            if tyh.time_differ(a, b) != 0:
                # print(tyh.time_differ(a, b))
                table_father.display(self, "时间：“时间”与《立案报告表》中记载的“案发时间”（" + b + "）不一致！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, text_temp[0], "“时间”与《立案报告表》中记载的“案发时间”（" + b + "）不一致")
            # Loc
            er = EntityRecognition()
            temp = er.get_identity_with_tag(self.contract_text, "LOC")
            loc_1 = ""
            for i in temp:
                loc_1 += i
            #print(loc_1)
            loc_2 = data.tabels_content["案发地点"]
            if loc_1 != loc_2:
                table_father.display(self, "地点： “地点”与《立案报告表》中记载的“案发地点”（"+loc_2+"）不一致！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, loc_1, "“地点”与《立案报告表》中记载的“案发地点”（"+loc_2+"）不一致")

    # 检查 案由、依据 是否与《立案报告表》和《案件处理审批表》中记载的 “案由”、“处罚依据” 一致。
    def check_reason1(self):  # reason_1 案由
        final_name = is_exist_cover("立案报告表", "撤销立案报告表", self.source_prefix)
        if not final_name:
            pass
            # table_father.display(self, "文书缺失：" + "《立案报告表》.docx不存在", "red")
        else:
            data_temp = DocxData(self.source_prefix + final_name)
            reason_1_1 = data_temp.tabels_content["案由"]
            reason_1_1 = reason_1_1.strip()
            if reason_1_1 not in self.contract_text:
                table_father.display(self, "案由：“案由”与《立案报告表》中记载的“案由”（"+reason_1_1+"）不一致！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "查获", "“案由”与《立案报告表》中记载的“案由”（"+reason_1_1+"）不一致!")

    def check_reason2(self):  # reason_2 依据
        if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
            pass
            # table_father.display(self, "文书缺失：" + "《案件处理审批表》.docx不存在", "red")
        else:
            data_temp1 = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
            # data_temp1 = DocxData(self.source_prefix + "案件处理审批表_.docx")
            reason2_1 = data_temp1.tabels_content["处罚依据"]
            # print(reason2_1)
            temp = re.search("依据(\s*\S+\s*)[的之]规定", self.contract_text)
            if temp:
                reason2_2 = temp.group(1).strip()
                #print(reason2_2)
                if reason2_2 not in reason2_1:
                    table_father.display(self, "处罚依据：“依据”与《案件处理审批表》中记载的“处罚依据”（"+reason2_1+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, reason2_2, "“依据”与《案件处理审批表》中记载的“处罚依据”（"+reason2_1+"）不一致!")
            else:
                table_father.display(self, "处罚依据：请人工核查《案件处理审批表》中的“处罚依据”与《案件处理审批表》中的“处罚依据” 一致性！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "依据", "请人工核查《案件处理审批表》中的“处罚依据”与《案件处理审批表》中的“处罚依据” 一致性！")

    # 检查 决定应当与《案件处理审批表》中的“承办人意见”一致。
    def check_Opinion_and_Content(self):
        # Opinion 承办人意见
        if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
            pass
            # table_father.display(self, "× " + "《案件处理审批表》.docx不存在", "red")
        else:
            data = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
            opinion = data.tabels_content["承办人意见"]
            #print(opinion)
            temp = re.search(r"作出(\s*\S+\s*)决定", self.contract_text)
            #print(temp)
            if temp:
                data3_Opinion = temp.group(1).strip()
                if data3_Opinion not in opinion:
                    table_father.display(self, "处罚决定：”处罚决定“与《案件处理审批表》中记载的“承办人意见”（"+opinion+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, data3_Opinion, "“处罚决定”与《案件处理审批表》中记载的““承办人意见”（"+opinion+"）不一致")
            else:
                table_father.display(self, "处罚决定：请人工核查“处罚决定”与《案件处理审批表》中的“承办人意见” 一致性！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "依据", "请人工核查“处罚决定”与《案件处理审批表》中的“承办人意见” 一致性！")

        # 送达的内容
        final_name = is_exist_cover("行政处罚决定书", "当场行政处罚决定书", self.source_prefix)
        if not final_name:
            table_father.display(self, "文书缺失：" + "《行政处罚决定书》.docx不存在", "red")
        else:
            # data = tyh.file_exists_open(self.source_prefix, "行政处罚决定书_", DocxData)
            data = DocxData(self.source_prefix + final_name)
            temp = data.text.split("\n")
            number = temp[2].strip()
            # print(number)
            #form_1 = number + "《行政处罚决定书》"   # ？？？？
            #form_2 = "《行政处罚决定书》" + number
            temp_content_form = re.search("现将(.*?)予以公告送达", self.contract_text)
            content_form = temp_content_form.group(1).strip()
            # print(content_form)
            if not (number in content_form and '行政处罚决定书' in content_form):
                table_father.display(self, "送达的内容：“送达的内容”应为 “行政处罚决定书”+编号，而合同中该字段不符合规定！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, content_form, "“送达的内容”应为 “行政处罚决定书”+编号 的形式")

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
    if "责令改正通知书_.docx" in list:
        ioc = Table_40(my_prefix, my_prefix)
        contract_file_path = my_prefix + "送达公告_.docx"
        ioc.check(contract_file_path, "送达公告_.docx")
