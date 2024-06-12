# -*- coding:utf-8 -*-
# @Author: tanweijia
import win32com
from win32com.client import Dispatch, constants
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.ReadFile import DocxData
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
from yancaoRegularDemo.Resource.tools.utils import *

function_description_dict = {
    'check_Name': '抬头应当与《行政处罚决定书》中的当事人保持一致',
    'check_Behavior_Content': '行为的内容，应当于《行政处罚决定书》中第4项对行为的定性一致，该文书表述为“系XXXXXXXXXX的违法行为。”',
    'check_Clause': '违反的条款的内容，应当与《行政处罚决定书》中第4项中，“当事人违反了XXXXXX的规定一致”',
    'check_Date': '日期应当与《行政处罚决定书》时间相同或在其之后。',
    'check_Review_Record': '复查记录：日期应在第5项日期之后，复查结果应当为“已整改”或“未整改”，如出现“未整改”即预警。',
    'check_Seal_and_itsDate': '复查日期应与复查记录的日期相同或之后。',
}

#my_prefix = r"D:\烟草项目\tobacco-xiejunyu\tobacco-xiejunyu\yancaoRegularDemo\副本"
# prefix = r'D:\烟草项目\tobacco-8.7\yancaoRegularDemo\副本'
# 责令改正通知书
class Table_41(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prefix = source_prefix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        # self.doc = self.mw.Documents.Open(self.my_prefix + "责令改正通知书_.docx")
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.check_Name()",
            "self.check_Behavior_Content()",
            "self.check_Clause()",
            "self.check_Seal()",
            "self.check_Date()",
            "self.check_Review_Record()",
            "self.check_Reviewer()",
            "self.check_Seal_and_itsDate()"
        ]
    # 检查 抬头 应当与《行政处罚决定书》中的当事人保持一致
    def check_Name(self):
        if not tyh.file_exists(self.source_prefix, "责令改正通知书"):
            table_father.display(self, "文书缺失：" + "《责令改正通知书》.docx不存在", "red")
        else:
            if not tyh.file_exists(self.source_prefix, "行政处罚决定书"):
                table_father.display(self, "文书缺失：" + "《行政处罚决定书》.docx不存在", "red")
            else:
                data2 = tyh.file_exists_open(self.source_prefix, "行政处罚决定书", DocxData)
                temp = self.contract_text.split()
                # print(temp)
                name1 = temp[3].rstrip("：")
                # print(name1)
                temp = re.search("当事人：(.*?)\n", data2.text)
                name2 = temp.group(1).strip()
                # print(name2)
                if name1 != name2:
                    table_father.display(self,"抬头：“文件抬头”与《行政处罚决定书》中的”当事人“（"+name2+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, name1, "该抬头与《行政处罚决定书》中记载的“当事人”（"+name2+"）不一致")

    # 检查 行为的内容，应当于《行政处罚决定书》中第4项对行为的定性一致，该文书表述为“系XXXXXXXXXX的违法行为。”
    def check_Behavior_Content(self):
        if not tyh.file_exists(self.source_prefix, "责令改正通知书"):
            pass
            # table_father.display(self, "× " + "《责令改正通知书》.docx不存在", "red")
        else:
            if not tyh.file_exists(self.source_prefix, "行政处罚决定书"):
                pass
                # table_father.display(self, "× " + "《行政处罚决定书》.docx不存在", "red")
            else:
                temp = re.search("你（单位）(.*?)的行为，", self.contract_text)
                if temp:
                    behavior1 = temp.group(1).strip()
                    #print(behavior1)
                    data2 = tyh.file_exists_open(self.source_prefix, "行政处罚决定书", DocxData)
                    temp = re.search("案由：(.*?)\n", data2.text)
                    behavior2 = temp.group(1).strip()
                    #print(behavior2)
                    if behavior1 != behavior2:
                        table_father.display(self, "行为的内容：”行为的内容“与《行政处罚决定书》中第4项对行为的定性（"+behavior2+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, behavior1, "”行为的内容“与《行政处罚决定书》中第4项对行为的定性（"+behavior2+"）不一致")
                else:
                    table_father.display(self, "行为的内容：请人工核查《责令改正通知书》中当事人行为的内容与《行政处罚决定书》中内容的一致性！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "", "请人工核查《责令改正通知书》中当事人行为的内容与《行政处罚决定书》中内容的一致性")

    # 检查 违反的条款的内容，应当与《行政处罚决定书》中第4项中，当事人违反了XXXXXX的规定 是否一致
    def check_Clause(self):
        if not tyh.file_exists(self.source_prefix, "责令改正通知书"):
            pass
            # table_father.display(self, "× " + "《责令改正通知书》.docx不存在", "red")
        else:
            if not tyh.file_exists(self.source_prefix, "行政处罚决定书"):
                pass
                # table_father.display(self, "× " + "《行政处罚决定书》.docx不存在", "red")
            else:
                temp = re.search("违反了(.*?)的规定", self.contract_text)
                if temp:
                    clause1 = temp.group(1).strip()
                    # print(seal1)
                    data2 = tyh.file_exists_open(self.source_prefix, "行政处罚决定书", DocxData)
                    temp = re.search("违反了(.*?)的规定。", data2.text)
                    if temp:
                        clause2 = temp.group(1).strip()
                        # print(seal2)
                        if clause1 != clause2:
                            table_father.display(self,"违反条款的内容：”违反条款的内容“与《行政处罚决定书》中第4项中‘当事人违反了XXXXXX的规定’不一致！", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, clause1, "”违反条款的内容“与《行政处罚决定书》中第4项中‘违反了"+clause2+"的规定’不一致")


    # 检查 落款处应加盖印章        ？ 该需求不做
    def check_Seal(self):
        pass

    # 检查 日期应当与《行政处罚决定书》时间相同或在其之后。
    def check_Date(self):
        if not tyh.file_exists(self.source_prefix, "责令改正通知书"):
            pass
            # table_father.display(self, "× " + "《责令改正通知书》.docx不存在", "red")
        else:
            if not tyh.file_exists(self.source_prefix, "行政处罚决定书"):
                pass
                # table_father.display(self, "× " + "《行政处罚决定书》.docx不存在", "red")
            else:
                temp = re.findall("(.*?)年(.*?)月(.*?)日\n", self.contract_text)
                # print(temp)
                Date_1 = str(temp[0][0]) + "年" + str(temp[0][1]) + "月" + str(temp[0][2]) + "日"
                real_date_1 = chinese_to_date(Date_1)
                # print(real_date_1)
                data2 = tyh.file_exists_open(self.source_prefix, "行政处罚决定书", DocxData)
                temp = re.search("(.*?)年(.*?)月(.*?)日\n", data2.text)
                # print(temp.group())
                Date_2 = temp.group().strip()
                real_date_2 = chinese_to_date(Date_2)
                # print(real_date_2)
                days = tyh.time_differ(real_date_1, real_date_2)
                if days >= 0:
                    pass
                else:
                    table_father.display(self,"日期：“日期“应与《行政处罚决定书》时间（"+Date_2+"）相同或在其之后！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, Date_1, "“日期“应与《行政处罚决定书》时间（"+Date_2+"）相同或在其之后！")

    # 检查 复查记录：日期应在第5项日期之后，复查结果应当为“已整改”或“未整改”，如出现“未整改”即预警。
    def check_Review_Record(self):
        if not tyh.file_exists(self.source_prefix, "责令改正通知书"):
            pass
            # table_father.display(self, "× " + "《责令改正通知书》.docx不存在", "red")
        else:
            temp = re.findall("复查记录：(.*?)我局", self.contract_text)
            mark = str(temp[0])
            review_date = tyh.get_strtime(temp[0].strip())
            # print(review_date)
            temp = re.findall("(.*?)年(.*?)月(.*?)日\n", self.contract_text)
            # print(temp)
            Date_1 = str(temp[0][0]) + "年" + str(temp[0][1]) + "月" + str(temp[0][2]) + "日"
            first_date = chinese_to_date(Date_1)
            # print(first_date)
            days = tyh.time_differ(review_date, first_date)
            if days > 0:
                pass
            else:
                table_father.display(self,"复查日期：”复查日期“应在责令日期（"+Date_1+"）之后，文件中日期有误！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, mark, "”复查日期“应该在责令日期（"+Date_1+"）之后，此处有误！")
            # 整改结果
            temp = re.search("复查结果如下：(.*?)。", self.contract_text)
            # print(temp.group(1))
            review_result = temp.group(1)
            if "已整改" in review_result or "已改正" in review_result:
                pass
            elif "未整改" in review_result:
                table_father.display(self,"复查结果：“复查结果”显示'未整改'！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, review_result, "复查结果 显示'未整改'!")
            else:
                table_father.display(self,"复查结果：“复查结果”撰写不合要求，应当为“已整改”或”未整改”", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, review_result, "复查结果应当为'已整改'或'未整改'")

    # 检查 复查人：签名完整，应当为两名以上执法人员，并记录执法证号。
    # ？ 暂未实现识别签名图片的方法，此处先实现识别执法证号
    def check_Reviewer(self):
        pass
    #     if not tyh.file_exists(self.source_prefix, "责令改正通知书_"):
    #         table_father.display(self, "× " + "《责令改正通知书》.docx不存在", "red")
    #     else:
    #         temp = re.findall("复查人（签名）：\s*(\d+)\n\s*(\d+)\n", self.contract_text)
    #         #print(temp)
    #         if len(temp) < 2:
    #             table_father.display(self, "× 复查人 少于两名以上执法人员！", "red")
    #             table_father.display(self, "复查人的执法人员编号分别为：" + str(temp), "green")
    #             tyh.addRemarkInDoc(self.mw, self.doc, "复查人（签名）：", "撰写不合要求，应为两名以上执法人员")

    # 检查 复查记录落款应加盖印章，日期应与复查记录的日期相同或之后。
    # 其中 暂时无法识别 印章图片
    def check_Seal_and_itsDate(self):
    # 检查印章 （pass）

    # 检查 日期应与复查记录的日期相同或之后
        if not tyh.file_exists(self.source_prefix, "责令改正通知书"):
            pass
            # table_father.display(self, "× " + "《责令改正通知书》.docx不存在", "red")
        else:
            temp = re.findall("(.*?)年(.*?)月(.*?)日\n", self.contract_text)
            # print(temp)
            chinese_Date = str(temp[1][0]) + "年" + str(temp[1][1]) + "月" + str(temp[1][2]) + "日"
            last_date = chinese_to_date(chinese_Date)
            # print(last_date)
            temp = re.findall("复查记录：(.*?)我局", self.contract_text)
            if temp:
                review_date = tyh.get_strtime(temp[0])
                # print(review_date)
                days = tyh.time_differ(last_date, review_date)
                if days >= 0:
                    pass
                else:
                    table_father.display(self, "落款日期：落款日期应与复查记录的日期（"+temp[0]+"）相同或在其之后，文件中不合要求！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, chinese_Date, "落款日期应与复查记录的日期（"+temp[0]+"）相同或之后！")

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
        ioc = Table_41(my_prefix, my_prefix)
        contract_file_path = my_prefix + "责令改正通知书_.docx"
        ioc.check(contract_file_path, "责令改正通知书_.docx")