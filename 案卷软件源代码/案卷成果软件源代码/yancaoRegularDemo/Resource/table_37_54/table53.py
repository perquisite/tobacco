# -*- coding:utf-8 -*-
# @Author: tanweijia
from os.path import dirname, abspath

import win32com
from win32com.client import Dispatch, constants
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.ReadFile import DocxData
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
from yancaoRegularDemo.Resource.tools.utils import *
from yancaoRegularDemo.Resource.tools.simple_content import Simple_Content
from docx import Document

function_description_dict = {
    'check_File_Name': '“文件名称”：可能为空，举报记录表、立案（不予立案）报告表、延长立案审批表、延长立案告知书、指定管辖通知书、检查（勘验）笔录、证据先行登记保存批准书、证据先行登记保存通知书、抽样取证物品清单、询问笔录、涉案烟草专卖品核价表、证据复制（提取）单、公告、损耗费用审批表、案件移送函、案件移送回执、移送财物清单、协助调查函、撤销立案报告表、案件调查终结报告、延长调查终结审批表、延长调查期限告知书、先行登记保存证据处理通知书、涉案物品返还清单、行政处罚事先告知书、陈述申辩记录、听证告知书、听证通知书、不予受理听证通知书、听证公告、听证笔录、听证报告、案件集体讨论记录、案件处理审批表、当场行政处罚决定书、行政处罚决定书、行政处理决定书、不予行政处罚决定书、送达回证、送达公告、责令改正通知书、责令整顿通知书、整顿终结通知书、违法物品销毁记录表、罚没变价处理审批表、罚没物品移交单、加处罚款决定书、强制执行申请书、延期缴款审批表、结案报告表、卷宗封面、卷宗目录、卷内备考表',
    'check_Page_Order': '“页次”：可能为空，应该连续',
}

# 卷宗目录
class Table_53(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prefix = source_prefix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("word.Application")
        # self.doc = self.mw.Documents.Open(self.my_prefix + "卷宗目录_.docx")
        self.contract_text = None
        self.contract_tables_content = None
        self.file_name_real = ""

        self.all_to_check = [
            "self.check_File_Name()",
            "self.check_Page_Order()"
            # "self.check_Page_Num()"
        ]

    # 检查 文件名称
    def check_File_Name(self):
        if not tyh.file_exists(self.source_prefix, "卷宗目录"):
            table_father.display(self, "× 《卷宗目录》.docx不存在", "red")
        else:
            Name_List = ['', '举报记录表', '立案报告表', '不予立案报告表', '延长立案审批表', '延长立案告知书', '指定管辖通知书', '检查（勘验）笔录', '证据先行登记保存批准书',
                         '证据先行登记保存通知书',
                         '抽样取证物品清单', '询问笔录', '涉案烟草专卖品核价表', '证据复制（提取）单', '公告', '损耗费用审批表', '案件移送函', '案件移送回执', '移送财物清单',
                         '协助调查函',
                         '撤销立案报告表', '案件调查终结报告', '延长调查终结审批表', '延长调查期限告知书', '先行登记保存证据处理通知书', '涉案物品返还清单', '行政处罚事先告知书',
                         '陈述申辩记录',
                         '听证告知书', '听证通知书', '不予受理听证通知书', '听证公告', '听证笔录', '听证报告', '案件集体讨论记录', '案件处理审批表', '当场行政处罚决定书',
                         '行政处罚决定书', '行政处理决定书',
                         '不予行政处罚决定书', '送达回证', '送达公告', '责令改正通知书', '责令整顿通知书', '整顿终结通知书', '违法物品销毁记录表', '罚没变价处理审批表',
                         '罚没物品移交单', '加处罚款决定书',
                         '强制执行申请书', '延期缴款审批表', '结案报告表', '卷宗封面', '卷宗目录', '卷内备考表']

            for file_name in self.contract_tables_content['题名']:
                # print(file_name)
                if file_name not in Name_List:
                    table_father.display(self, "文件名称：“题名”中的文件名称“" + file_name + "”不存在或不规范！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, file_name, "该文件名称不存在或不规范！")

    # 检查 页次 是否连续
    def check_Page_Order(self):
        def Front(x):
            a = x.split('-')[0]
            return int(a)

        def Tail(x):
            a = x.split('-')[1]
            return int(a)
        file_path = self.source_prefix + '\\'+ self.file_name_real
        #print(file_path)
        doc = Document(file_path)
        tb = doc.tables[0]
        tb.columns
        column_cells = tb.columns[5].cells  # 对 文件名称 的那一列
        page_List = []
        for i in range(1, len(tb.rows)):
            page_List.append(column_cells[i].text.strip())
        origin_List = page_List
        #print(origin_List)
        # 去除空元素和“/”
        page_List = [i for i in page_List if (i not in ["\\", "/"] and (len(str(i))) != 0)]
        #print(page_List)
        final_list = []
        for j in range(len(page_List)):
            if not ('-' in page_List[j]):
                final_list.append(page_List[j] + '-' + page_List[j])
            else:
                final_list.append(page_List[j])
        #print(final_list)
        max = len(page_List) - 1
        k = 1
        while k <= max:
            F = Front(final_list[k])
            T = Tail(final_list[k - 1])
            if (F - T == 1):
                serial = True  # 此处连续
            else:
                serial = False
            if serial:
                pass  # 连续
                # print("连续")
            else:
                #print("不连续")
                table_father.display(self, "× 页次不连续！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, page_List[k], "该处页次不连续！")
            k = k + 1

    # 检查 备注
    # def check_Page_Num(self):
    #     data = DocxData(self.source_prefix + "卷宗目录_.docx")
    #     temp1 = re.search("共计(.*?)页", data.text)
    #     #page_Num = temp1.group(1).strip()
    #     temp2 = re.search("附证物(.*?)袋", data.text)
    #     #evidence_Num = temp2.group(1).strip()
    #     if temp1 == None:
    #         table_father.display(self,"× 共计页数 为空！", "red")
    #         tyh.addRemarkInDoc(self.mw, self.doc, "共计", "共计页数不应为空！！")
    #     if temp2 == None:
    #         table_father.display(self,"× 附证物袋数 为空！", "red")
    #         tyh.addRemarkInDoc(self.mw, self.doc, "附证物", "附证物袋数不应为空！！")

    def check(self, contract_file_path, file_name_real):
        self.file_name_real = file_name_real
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
    source_prefix = "C:\\Users\\twj\Desktop\\test\\"
    # print(my_prefix)
    ioc = Table_53(source_prefix, source_prefix)
    contract_file_path = source_prefix + "卷宗目录_.docx"
    # print(contract_file_path)
    ioc.check(contract_file_path, "卷宗目录_.docx")
