# -*- coding:utf-8 -*-
# @Author: tanweijia
import time
from os.path import dirname, abspath

import win32com
from win32com.client import Dispatch, constants
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.ReadFile import DocxData
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os

from yancaoRegularDemo.Resource.tools.tanweijia_function import is_exist_cover
from yancaoRegularDemo.Resource.tools.utils import *
from yancaoRegularDemo.Resource.tools.simple_content import Simple_Content
from docx import Document

function_description_dict = {
    'check_Number': '“年度  第  号”：不为空，形式为“xxxx年度x烟第xxx号”。',
    'check_Reason': '“案由”： 应与案卷中《案件处理审批表》、《案件集体讨论记录》、《陈述申辩记录》、《结案报告》中“案由一致”。',
    'check_Client': '“当事人”：不为空，与《卷宗封面》、《立案报告表》、《抽样取证物品清单》、《涉案烟草专卖品核价表》、《调查总结报告》、《延长案件调查终结审批表》、《案件处理审批表》、《先行登记保存证据处理通知书》、《行政处罚事先告知书》、《听证告知书》、《当场行政处罚决定书》、《行政处罚决定书》、《违法物品销毁记录表》、《结案报告》中“当事人”一致',
    'check_3_Date_and_StorageTime': '1、“立案日期”： 不为空，填写卷宗“立案报告表”中领导批准立案的日期。 与《涉案烟草专卖品核价表》、《调查总结报告》、《延长案件调查终结审批表》、《案件处理审批表》、《卷宗封面》的立案日期一致；2、“结案日期”： 不为空，栏应填写卷宗“结案报告表”中领导批准结案的日期。；3、“归档日期”： 不为空，在“结案日期”之后；“保存期限”：不为空，永久',
    'check_Approver_and_Undertaker': '1、“审批人”： 不为空；2、“承办人”： 不为空，与案卷中承办人一致',
    'check_Result_and_PageNum': '“处理结果”：不为空，与《行政处罚决定书》中的处罚决定一致',
}

# 卷宗封面
class Table_52(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prefix = source_prefix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.contract_text = None
        self.contract_tables_content = None
        self.file_name_real = ''

        self.all_to_check = [
            "self.check_Number()",
            "self.check_Reason()",
            "self.check_Client()",
            "self.check_3_Date_and_StorageTime()",  #三个日期和保存期限
            "self.check_Approver_and_Undertaker()",  #审批人和承办人
            "self.check_Result_and_PageNum()"    #“处理结果”和 共计几页
        ]

    # 检查 “年度  第  号”：不为空，形式为“xxxx年度x烟第xxx号”。
    def check_Number(self):
        if not tyh.file_exists(self.source_prefix, "卷宗封面"):
            table_father.display(self, "× " + "《卷宗封面》.docx不存在", "red")
        else:
            file_path = self.source_prefix + self.file_name_real
            # print(file_path)
            doc = Document(file_path)
            tb = doc.tables[0]
            tb.rows
            row_cells = tb.rows[1].cells
            Number_text = row_cells[0].text
            #print(Number_text)
            #result = re.match("(\d*?)年度(.*?)第(\d*?)号", Number_text)
            #print(result)
            text_1 = re.match(r"(\d*?)年度", Number_text)
            # print(text_1)
            text_2 = re.search(r"年度(\w*?)第", Number_text)
            #print(text_2)
            text_3 = re.search(r"第(\d*?)号", Number_text)
            # print(text_3.group(1))
            sp = Simple_Content()
            # 检查text_1: 年度
            if not text_1:   # 没有填写 年度，或 填的不是数字\d，包括空格
                #print("空1")
                table_father.display(self,"年度：没有填写“年度”，或填写的是非数字！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "年度", "“年度”填写不合要求")
            elif sp.is_null(text_1.group(1)): # "年度"前没有填写任何东西
                #print("空2")
                table_father.display(self,"年度：没有填写“年度”！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "年度", "“年度”填写不合要求")
            else:  # 填写了年度，进一步检查是不是数字 是不是4位的年份
                #print("不空")
                if len(text_1.group(1)) != 4:
                    #print("年度格式有误！")
                    table_father.display(self,"年度：”年度“填写格式有误 ！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "年度", "“年度”填写不合要求")

            # 检查text_2: x烟
            if not text_2: # 填写了非“\w”，如 空格
                #print("空1")
                table_father.display(self,"年度第X号：'xxxx年度x烟第xxx号'中第二项填写不合要求 ！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "第", "年度之后、号之前内容填写不合要求")
            elif sp.is_null(text_2.group(1)): # 没有填写
                #print("空2")
                table_father.display(self,"年度第X号：没有填写'xxxx年度x烟第xxx号'中第二项！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "第", "年度之后、号之前内容填写不合要求")
            else:
                if len(text_2.group(1)) == 2 and "烟" == text_2.group(1)[1]:  # 判断是否为 X烟 的形式
                    pass
                    #print("符合要求")
                else:
                    #print("不符合要求")
                    table_father.display(self,"年度第X号：没有填写 'xxxx年度x烟第xxx号' 中第二项x烟填写不合要求 ！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "第", "年度之后、号之前内容填写不合要求")

            # 检查text_3: 第xxx号
            if not text_3:   # 没有填写，或 填的不是数字\d(包括空格)
                #print("空1")
                table_father.display(self,"年度第X号：没有填写第几号，或填写的是非数字！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "号", "填写不合要求")
            elif sp.is_null(text_3.group(1)): # 没有填写任何东西
                #print("空2")
                table_father.display(self,"年度第X号：没有填写第几号 ！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "号", "填写不合要求")
            else:  # 填写了，进一步检查是不是数字 是不是最多3位的号码数
                if 0 < len(text_3.group(1)) <= 3:
                    pass
                    #print("合格！")
                else:
                    table_father.display(self,"年度第X号：第几号填写格式有误 ！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "号", "填写不合要求")

    # 检查 案由
    def check_Reason(self):
        if not tyh.file_exists(self.source_prefix, "卷宗封面"):
            pass
            # table_father.display(self, "× " + "《卷宗封面》.docx不存在", "red")
        else:
            data0 = tyh.file_exists_open(self.source_prefix, "卷宗封面", DocxData)
            Reason = data0.tabels_content["案由"]
            #print(Reason)
            # 1 与 《行政处罚决定书》中的 案由 比较
            if not tyh.file_exists(self.source_prefix, "行政处罚决定书"):
                table_father.display(self, "文书缺失：" + "《行政处罚决定书》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "行政处罚决定书", DocxData)
                temp = re.findall(r"案由\s*：(.*?)\n", data.text)
                # print(temp)
                if temp:
                    reason_1 = temp[0].strip()
                    if Reason != reason_1:
                        table_father.display(self,"案由：”案由“与《行政处罚决定书》中的“案由”（"+reason_1+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "案   由", "”案由“与《行政处罚决定书》中的“案由”（"+reason_1+"）不一致")
                else:
                    table_father.display(self, "案由：未从《行政处罚决定书》中提取到“案由”，因此无法对比，请人工核查！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "案   由", "未从《行政处罚决定书》中提取到“案由”，因此无法对比，请人工核查！")

            # 2 与 《案件处理审批表》中的 案由 比较
            if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
                table_father.display(self, "文书缺失：" + "《案件处理审批表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
                reason_2 = data.tabels_content["案由"]
                #print(reason_2)
                if Reason != reason_2:
                    table_father.display(self,"案由：”案由“与《案件处理审批表》中的 案由（"+reason_2+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "案   由", "”案由“与《案件处理审批表》中的”案由“（"+reason_2+"）不一致")

            # 3 与 《延期（分期）缴纳罚款审批表》中的 案由 比较          !!!没有此表模板    此表不做
            # if os.path.exists(self.source_prefix + "延期缴纳罚款审批表_.docx") == 1:
            #     data = DocxData(self.source_prefix + "延期缴纳罚款审批表_.docx")
            #     reason_3 = data.tabels_content["案由"]
            #     if Reason != reason_3:
            #         table_father.display(self,"× 案由 与《延期缴纳罚款审批表》中的 案由 不一致！", "red")
            #         tyh.addRemarkInDoc(self.mw, self.doc, "案   由", "与《延期缴纳罚款审批表》中的 案由 不一致")
            #
            # elif os.path.exists(self.source_prefix + "分期缴纳罚款审批表_.docx") == 1:
            #     data = DocxData(self.source_prefix + "分期缴纳罚款审批表_.docx")
            #     reason_3 = data.tabels_content["案由"]
            #     if Reason != reason_3:
            #         table_father.display(self,"× 案由 与《分期缴纳罚款审批表》中的 案由 不一致！", "red")
            #         tyh.addRemarkInDoc(self.mw, self.doc, "案   由", "与《分期缴纳罚款审批表》中的 案由 不一致")
            # else:
            #     table_father.display(self,"× 《延期（分期）缴纳罚款审批表》.docx不存在", "red")


            # 4 与 《案件集体讨论记录》中的 案由 比较
            if not tyh.file_exists(self.source_prefix, "案件集体讨论记录"):
                table_father.display(self, "文书缺失：" + "《案件集体讨论记录》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "案件集体讨论记录", DocxData)
                temp = re.findall("案由：(.*?)\n", data.text)
                reason_4 = temp[0].strip()
                #print(reason_4)
                if Reason != reason_4:
                    table_father.display(self,"案由：”案由“与《案件集体讨论记录》中的”案由“（"+reason_4+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "案   由", "”案由“与《案件集体讨论记录》中的”案由“（"+reason_4+"）不一致")

            # 5 与 《陈述申辩记录》中的 案由 比较
            if not tyh.file_exists(self.source_prefix, "陈述申辩记录"):
                table_father.display(self, "文书缺失：" + "《陈述申辩记录》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "陈述申辩记录", DocxData)
                temp = re.findall("案由：(.*?)\n", data.text)
                reason_5 = temp[0].strip()
                #print(reason_5)
                if Reason != reason_5:
                    table_father.display(self,"案由：”案由“与《陈述申辩记录》中的 案由（"+reason_5+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "案   由", "”案由“与《陈述申辩记录》中的”案由“（"+reason_5+"）不一致")

            # 6 与 《结案报告表》中的 案由 比较
            if not tyh.file_exists(self.source_prefix, "结案报告表"):
                table_father.display(self, "文书缺失：" + "《结案报告表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "结案报告表", DocxData)
                reason_6 = data.tabels_content["案由"]
                #print(reason_6)
                if Reason != reason_6:
                    table_father.display(self,"案由：”案由“与《结案报告表》中的”案由“（"+reason_6+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "案   由", "”案由“与《结案报告表》中的”案由“（"+reason_6+"）不一致")

    # 检查 当事人
    def check_Client(self):
        if not tyh.file_exists(self.source_prefix, "卷宗封面"):
            table_father.display(self, "文书缺失：" + "《卷宗封面》.docx不存在", "red")
        else:
            data0 = tyh.file_exists_open(self.source_prefix, "卷宗封面", DocxData)
            Client = data0.tabels_content["当事人"]
            sp = Simple_Content()
            if sp.is_null(Client):
                table_father.display(self,"当事人：× “当事人” 为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "不应为空！")
            else:
                # 1 与 《立案报告表》中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "立案报告表"):
                    table_father.display(self, "文书缺失：" + "《立案报告表》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "立案报告表", DocxData)
                    client_1 = data.tabels_content["当事人"]
                    #print(client_1)
                    if Client != client_1:
                        table_father.display(self,"当事人：× “当事人”与《立案报告表》中的“当事人”（"+client_1+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《立案报告表》中的“当事人”（"+client_1+"）不一致")

                # 2 与 《延长立案期限审批表》中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "延长立案期限审批表"):
                    table_father.display(self, "文书缺失：" + "《延长立案期限审批表》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "延长立案期限审批表", DocxData)
                    client_2 = data.tabels_content["当事人"]
                    #print(client_2)
                    if Client != client_2:
                        table_father.display(self,"当事人：× “当事人”与《延长立案期限审批表》中的“当事人”（"+client_2+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《延长立案期限审批表》中的“当事人”（"+client_2+"）不一致")

                # 3 与 《证据先行登记保存通知书》中的 “当事人”比较    (此表无文件，先跳过)
                if not tyh.file_exists(self.source_prefix, "先行登记保存证据处理通知书"):
                    table_father.display(self, "文书缺失：" + "《先行登记保存证据处理通知书》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "先行登记保存证据处理通知书", DocxData)
                    temp = data.text.split()
                    client_3 = temp[3].strip("：")
                    print("当事人" + client_3)
                    if Client != client_3:
                        table_father.display(self,"当事人：“当事人”与《先行登记保存证据处理通知书》中的“当事人”（"+client_3+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《先行登记保存证据处理通知书》中的“当事人”（"+client_3+"）不一致")

                # 4 与 《抽样取证物品清单》中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "抽样取证物品清单"):
                    table_father.display(self, "文书缺失：" + "《抽样取证物品清单》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "抽样取证物品清单", DocxData)
                    client_4 = data.tabels_content["当事人"]
                    #print(client_4)
                    if Client != client_4:
                        table_father.display(self,"当事人：“当事人”与《抽样取证物品清单》中的“当事人”（"+client_4+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《抽样取证物品清单》中的“当事人”（"+client_4+"）不一致")

                # 5 与 《涉案烟草专卖品核价表》中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "涉案烟草专卖品核价表"):
                    table_father.display(self, "文书缺失：" + "《涉案烟草专卖品核价表》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "涉案烟草专卖品核价表", DocxData)
                    temp = re.findall("当事人：(.*?)\n", data.text)
                    if temp:
                        client_5 = temp[0].strip()
                        #print(client_5)
                        if Client != client_5:
                            table_father.display(self,"当事人：“当事人”与《涉案烟草专卖品核价表》中的“当事人”（"+client_5+"）不一致！", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《涉案烟草专卖品核价表》中的“当事人”（"+client_5+"）不一致")
                    else:
                        table_father.display(self, "当事人：《涉案烟草专卖品核价表》中的没有提取到“当事人”，故无法和“当事人对比”！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "《涉案烟草专卖品核价表》中的没有提取到“当事人”，故无法和“当事人对比”！")

                # 6 与 《卷烟鉴别检验样品留样、损耗费用审批表》中的 “当事人”比较   (此表无文件，先跳过)
                flag = False
                if tyh.file_exists(self.source_prefix, "卷烟鉴别检验样品留样、损耗费用审批表"):
                    data = tyh.file_exists_open(self.source_prefix, "卷烟鉴别检验样品留样、损耗费用审批表", DocxData)
                    flag = True
                elif tyh.file_exists(self.source_prefix, "损耗费用审批表"):
                    data = tyh.file_exists_open(self.source_prefix, "损耗费用审批表", DocxData)
                    flag = True
                else:
                    table_father.display(self, "文书缺失：" + "《卷烟鉴别检验样品留样、损耗费用审批表》.docx不存在", "red")
                if flag:
                    client_6 = data.tabels_content["案件当事人"]
                    if client_6 != Client:
                        table_father.display(self, "当事人： “当事人”与《卷烟鉴别检验样品留样、损耗费用审批表》中的“当事人”（"+client_6+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《卷烟鉴别检验样品留样、损耗费用审批表》中的“当事人”（"+client_6+"）不一致")

                # 7 与 《调查总结报告》(《调查终结报告》)中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "调查终结报告"):
                    table_father.display(self, "文书缺失：" + "《调查终结报告》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "调查终结报告", DocxData)
                    client_7 = data.tabels_content["当事人"]
                   #print(client_7)
                    if Client != client_7:
                        table_father.display(self,"当事人：“当事人”与《调查终结报告》中的“当事人”（"+client_7+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《调查终结报告》中的“当事人”（"+client_7+"）不一致")

                # 8 与 《延长调查终结审批表》(《延长案件调查终结审批表》)中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "延长调查终结审批表"):
                    table_father.display(self, "文书缺失：" + "《延长调查终结审批表》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "延长调查终结审批表", DocxData)
                    client_8 = data.tabels_content["当事人"]
                    #print(client_8)
                    if Client != client_8:
                        table_father.display(self,"当事人：“当事人”与《延长调查终结审批表》中的“当事人”（"+client_8+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《延长调查终结审批表》中的“当事人”（"+client_8+"）不一致")

                # 9 与 《案件处理审批表》 中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
                    table_father.display(self, "文书缺失：" + "《案件处理审批表》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
                    client_9 = data.tabels_content["当事人"]
                    #print(client_9)
                    if Client != client_9:
                        table_father.display(self, "当事人：“当事人”与《案件处理审批表》中的“当事人”"+str(client_9)+"不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《案件处理审批表》中的“当事人”"+str(client_9)+"不一致")

                # 10 与 《先行登记保存证据处理通知书》 中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "先行登记保存证据处理通知书"):
                    table_father.display(self, "文书缺失：《先行登记保存证据处理通知书》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "先行登记保存证据处理通知书", DocxData)
                    temp = data.text.split("\n")
                    client_10 = temp[3].strip().rstrip("：").rstrip(":").strip()
                    #print(client_10)
                    if Client != client_10:
                        table_father.display(self,"当事人：“当事人”与《先行登记保存证据处理通知书》中的“当事人”（"+client_10+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《先行登记保存证据处理通知书》中的“当事人”（"+client_10+"）不一致")

                # 11 与 《行政处罚事先告知书》 中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "行政处罚事先告知书"):
                    table_father.display(self, "文书缺失：《行政处罚事先告知书》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "行政处罚事先告知书", DocxData)
                    temp = data.text.split("\n")
                    client_11 = temp[3].strip().rstrip("：").rstrip(":").strip()
                    #print(client_11)
                    if Client != client_11:
                        table_father.display(self,"当事人：“当事人”与《行政处罚事先告知书》中的“当事人”（"+client_11+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《行政处罚事先告知书》中的“当事人”（"+client_11+"）不一致")

                # 12 与 《听证告知书》(《听证告知》) 中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "听证告知"):
                    table_father.display(self, "文书缺失：《听证告知》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "听证告知", DocxData)
                    temp = data.text.split("\n")
                    client_12 = temp[3].strip().rstrip("：").rstrip(":").strip()
                    #print(client_12)
                    if Client != client_12:
                        table_father.display(self,"当事人：“当事人”与《听证告知》中的“当事人”（"+client_12+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《听证告知书》中的“当事人”（"+client_12+"）不一致")

                # 13 与 《听证通知书》中的 “当事人”比较   (此表无文件，先跳过)
                # if os.path.exists(self.source_prefix + "听证通知书_.docx") == 1:
                #     pass
                # else:
                #     table_father.display(self,"× 《听证通知书》.docx不存在", "red")

                # 14 与 《不予受理听证通知书》中的 “当事人”比较   (此表无文件，先跳过)
                # if os.path.exists(self.source_prefix + "不予受理听证通知书_.docx") == 1:
                #     pass
                # else:
                #     table_father.display(self,"× 《不予受理听证通知书》.docx不存在", "red")

                # 15 与 《听证笔录》中的 “当事人”比较   (此表无文件，先跳过)
                # if os.path.exists(self.source_prefix + "听证笔录_.docx") == 1:
                #     pass
                # else:
                #     table_father.display(self,"× 《听证笔录》.docx不存在", "red")

                # 16 与 《听证报告》中的 “当事人”比较   (此表无文件，先跳过)
                # if os.path.exists(self.source_prefix + "听证报告_.docx") == 1:
                #     pass
                # else:
                #     table_father.display(self,"× 《听证报告》.docx不存在", "red")

                # 17 与 《当场行政处罚决定书》中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "当场行政处罚决定书"):
                    table_father.display(self, "文书缺失：《当场行政处罚决定书》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "当场行政处罚决定书", DocxData)
                    temp = re.search("当事人名称（姓名）：(.*?)\n", data.text)
                    client_17 = temp.group(1).strip()
                    if Client != client_17:
                        table_father.display(self,"当事人：“当事人”与《当场行政处罚决定书》中的“当事人”（"+client_17+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《当场行政处罚决定书》中的“当事人”（"+client_17+"）不一致")

                # 18 与 《行政处罚决定书》中的 “当事人”比较
                file_name = not is_exist_cover("行政处罚决定书", "当场行政处罚决定书", self.source_prefix)
                if not file_name:
                    table_father.display(self, "文书缺失：《行政处罚决定书》.docx不存在", "red")
                else:
                    data = DocxData(self.source_prefix + file_name)
                    temp = re.findall("当事人：(.*?)\n", data.text)
                    client_18 = temp[0].strip()
                    #print(client_18)
                    if Client != client_18:
                        table_father.display(self,"当事人：“当事人”与《行政处罚决定书》中的“当事人”（"+client_18+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《行政处罚决定书》中的“当事人”（"+client_18+"）不一致")

                # 19 与 《违法物品销毁记录表》中的 “当事人”比较  (此表无文件，先跳过)
                if not tyh.file_exists(self.source_prefix, "违法物品销毁记录表"):
                    table_father.display(self, "文书缺失：《违法物品销毁记录表》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "违法物品销毁记录表", DocxData)
                    client_19 = data.tabels_content["当事人"]
                    if Client != client_19:
                        table_father.display(self, "当事人：“当事人”与《违法物品销毁记录表》中的“当事人”（"+client_19+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《违法物品销毁记录表》中的“当事人”（"+client_19+"）不一致")

                # 20 与 《加处罚款决定书》中的 “当事人”比较   (此表无文件，先跳过)
                # if os.path.exists(self.source_prefix + "加处罚款决定书_.docx") == 1:
                #     pass
                # else:
                #     table_father.display(self, "× 《加处罚款决定书》.docx不存在", "red")

                # 21 与 《延期（分期）缴纳罚款审批表》中的 “当事人”比较  (此表无文件，先跳过)
                # if os.path.exists(self.source_prefix + "延期缴纳罚款审批表_.docx") == 1:
                #     pass
                # elif os.path.exists(self.source_prefix + "分期缴纳罚款审批表_.docx") == 1:
                #     pass
                # else:
                #     table_father.display(self,"× 《延期（分期）缴纳罚款审批表》.docx不存在", "red")

                # 22 与 《结案报告》中的 “当事人”比较
                if not tyh.file_exists(self.source_prefix, "结案报告表"):
                    table_father.display(self, "文书缺失：《结案报告表》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "结案报告表", DocxData)
                    client_22 = data.tabels_content["当事人"]
                    #print(client_22)
                    if Client != client_22:
                        table_father.display(self,"当事人：“当事人”与《结案报告表》中的“当事人”（"+client_22+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "与《结案报告表》中的“当事人”（"+client_22+"）不一致")

    # 检查 三个日期（立案日期、结案日期、归档日期）和保存期限
    def check_3_Date_and_StorageTime(self):
        if tyh.file_exists(self.source_prefix, "卷宗封面"):
            data0 = self.contract_tables_content["立案日期"]
            # 立案日期                                                    (要求 有一半不明确)
            if data0 == "":
                table_father.display(self, "立案日期：“立案日期”不应为空！", "red")
            else:
                # print("立案\n")
                start_time = tyh.get_strtime(data0)
                if not tyh.file_exists(self.source_prefix, "立案报告表"):
                    pass
                    # table_father.display(self, "× 《立案报告表》.docx不存在", "red")
                else:
                    data1 = tyh.file_exists_open(self.source_prefix, "立案报告表", DocxData)
                    content = data1.tabels_content["负责人意见"]
                    temp = re.search(r"日期[：:](\s*\S+\s*年\s*\S+\s*月\s*\S+\s*日\s*)$", content)
                    if temp:
                        #print(temp)
                        leader_date = tyh.get_strtime(temp.group(1))
                        #print(leader_date)
                        #print(start_time)
                        if tyh.time_differ(start_time, leader_date):
                            table_father.display(self, "立案日期：“立案日期”与 “立案报告表”中领导批准立案的日期（"+temp.group(1)+"）不同！", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, "立 案 日 期", "与 “立案报告表”中 领导批准立案的日期（"+temp.group(1)+"）不一致")

                if not tyh.file_exists(self.source_prefix, "涉案烟草专卖品核价表"):
                    table_father.display(self, "文书缺失：《涉案烟草专卖品核价表》.docx不存在", "red")
                else:
                    data2 = tyh.file_exists_open(self.source_prefix, "涉案烟草专卖品核价表", DocxData)
                    temp = re.search("于(.*?)查获的", data2.text)
                    if temp:
                        # Date_2 = temp.group(1)
                        Date_to_Check = tyh.get_strtime(temp.group(1))
                        if tyh.time_differ(start_time, Date_to_Check):
                            table_father.display(self,"立案日期：”立案日期“与 “涉案烟草专卖品核价表”中的“立案日期”（"+temp.group(1)+"）不同！", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, "立 案 日 期", "与“涉案烟草专卖品核价表”中“立案日期”（"+temp.group(1) + "）不一致")

                if not tyh.file_exists(self.source_prefix, "调查终结报告"):
                    pass
                    # table_father.display(self, "× 《调查终结报告》.docx不存在", "red")
                else:
                    data3 = tyh.file_exists_open(self.source_prefix, "调查终结报告", DocxData)
                    Date_3 = data3.tabels_content["立案日期"]
                    #print(Date_3)
                    Date_to_Check = tyh.get_strtime(Date_3)
                    #print(Date_to_Check)
                    if tyh.time_differ(start_time, Date_to_Check):
                        table_father.display(self,"立案日期：”立案日期“与 “调查终结报告”中的“立案日期”（"+Date_3+"）不同！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "立 案 日 期", "与 “调查终结报告”中“立案日期”（"+Date_3+"）不一致")

                if not tyh.file_exists(self.source_prefix, "延长案件调查终结审批表"):
                    table_father.display(self, "文书缺失：《延长案件调查终结审批表》.docx不存在", "red")
                else:
                    data4 = tyh.file_exists_open(self.source_prefix, "延长案件调查终结审批表", DocxData)
                    Date_4 = data4.tabels_content["立案日期"]
                    #print(Date_4)
                    Date_to_Check = tyh.get_strtime(Date_4)
                    #print(Date_to_Check)
                    if tyh.time_differ(start_time, Date_to_Check):
                        table_father.display(self, "立案日期：”立案日期“与 “延长案件调查终结审批表”中的“立案日期”（"+Date_4+"）不同！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "立 案 日 期", "与 “延长案件调查终结审批表”中“立案日期”（"+Date_4+"）不一致")

                if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
                    table_father.display(self, "文书缺失：《案件处理审批表》.docx不存在", "red")
                else:
                    data5 = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
                    Date_5 = data5.tabels_content["立案日期"]
                    #print(Date_5)
                    Date_to_Check = tyh.get_strtime(Date_5)
                    #print(Date_to_Check)
                    if tyh.time_differ(start_time, Date_to_Check):
                        table_father.display(self,"立案日期：“立案日期”与 “案件处理审批表”中的 立案日期（"+Date_5+"）不同！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "立 案 日 期", "与 “案件处理审批表”中 “立案日期”（"+Date_5+"）不一致")

                if not tyh.file_exists(self.source_prefix, "卷宗封面"):
                    table_father.display(self, "文书缺失：《卷宗封面》.docx不存在", "red")
                else:
                    data6 = tyh.file_exists_open(self.source_prefix, "卷宗封面", DocxData)
                    Date_6 = data6.tabels_content["立案日期"]
                    #print(Date_6)
                    Date_to_Check = tyh.get_strtime(Date_6)
                    #print(Date_to_Check)
                    if tyh.time_differ(start_time, Date_to_Check):
                        table_father.display(self, "立案日期：“立案日期”与 “卷宗封面”中的 “立案日期”（"+Date_6+"）不同！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "立 案 日 期", "与 “卷宗封面”中 “立案日期”（"+Date_6+"）不一致")

            # 结案日期
            data0 = self.contract_tables_content["结案日期"]
            if data0 == "":
                table_father.display(self, "结案日期：“结案日期”不应为空！", "red")
            else:
                end_time = tyh.get_strtime(data0)
                if not tyh.file_exists(self.source_prefix, "结案报告表"):
                    table_father.display(self, "文书缺失：《结案报告表》.docx不存在，无法对比结案日期", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "结 案 日 期", "《结案报告表》.docx不存在，无法对比结案日期")
                else:
                    data2 = tyh.file_exists_open(self.source_prefix, "结案报告表", DocxData)
                    content = data2.tabels_content["负责人意见"]
                    temp = re.search(r"日期[：:](\s*\S+\s*年\s*\S+\s*月\s*\S+\s*日\s*)$", content)
                    if temp:
                        leader_date = tyh.get_strtime(temp.group(1))
                        if tyh.time_differ(end_time, leader_date):
                            table_father.display(self, "结案日期：“结案日期”与 “结案报告表”中领导批准结案的日期（"+temp.group(1)+"）不同！", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, "结 案 日 期", "与 “立案报告表”中领导批准立案的日期（"+temp.group(1)+"）不一致")
                    else:
                        table_father.display(self, "结案日期：《结案报告表》中没有检索到结案日期，请人工核查！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "结 案 日 期", "《结案报告表》中没有检索到结案日期，请人工核查！")
            # 归档日期
            data0 = self.contract_tables_content["归档日期"]
            if data0 == "":
                table_father.display(self, "归档日期：“归档日期”不应为空！", "red")
            else:
                store_time = tyh.get_strtime(data0)
                #print(store_time)
                #print(end_time)
                if tyh.time_differ(store_time, end_time) <= 0:
                    table_father.display(self, "归档日期：“归档日期”应在“结案日期”（"+end_time+"）之后，表中有误！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "归 档 日 期", "“归档日期”应在 结案日期（"+end_time+"）之后，表中有误！")

            # 保存期限
            data0 = self.contract_tables_content["保存期限"]
            if data0 == "":
                table_father.display(self, "保存期限：“保存期限”不应为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "保 存 期 限", "“保存期限”不应为空！")
            else:
                str1 = data0.replace(" ", "")
                #print(str1)
                if not "永久" in str1:
                    table_father.display(self,"保存期限：“保存期限”应该为 永久！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "保 存 期 限", "“保存期限”应该为 永久！")

    # 检查 审批人 和 承办人
    def check_Approver_and_Undertaker(self):
        if tyh.file_exists(self.source_prefix, "卷宗封面"):
            # 审批人不为空
            sc = Simple_Content()
            approver = self.contract_tables_content['审批人']
            #print(approver)
            if sc.is_null(approver):
                table_father.display(self, "审批人：审批人为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "审批人", "审批人为空！！")
            # 承办人 与案卷中承办人一致 （从结案报告表中提取）
            temp = self.contract_tables_content['承办人']
            #print(temp)
            if sc.is_null(temp):
                table_father.display(self, "承办人：承办人为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "承办人", "承办人为空！！")
            else:
                if tyh.file_exists(self.source_prefix, "结案报告表"):
                    undertaker_list = []
                    if " " in temp:
                        raw = temp.split(" ")
                        undertaker_list = [i for i in raw if not i==""]
                    elif "," in temp:
                        undertaker_list = temp.split(",")
                    elif "，" in temp:
                        undertaker_list = temp.split("，")
                    elif "、" in temp:
                        undertaker_list = temp.split("，")
                    else:
                        table_father.display(self, "承办人：承办人填写不规范！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "承办人", "承办人填写不规范！")
                    if undertaker_list:
                        pass
                    else:
                        data = tyh.file_exists_open(self.source_prefix, "结案报告表", DocxData)
                        compare = data.tabels_content['调查人']
                        #print(compare)
                        for item in undertaker_list:
                            if not item in compare:
                                table_father.display(self, "承办人：承办人与案卷的承办人（"+compare+"）不一致！请人工检查", "red")
                                tyh.addRemarkInDoc(self.mw, self.doc, "承办人", "承办人与案卷的承办人（"+compare+"）不一致！请人工检查")
                                break
                else:
                    table_father.display(self, "承办人：案卷文件夹内不含《结案报告表》，因此无法对比“承办人”，请人工核验！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "承办人", "案卷文件夹内不含《结案报告表》，因此无法对比“承办人”，请人工核验！")

    # 检查  “处理结果”和 共计几页
    def check_Result_and_PageNum(self):
        # “处理结果”： 不为空，与《行政处罚决定书》中的处罚决定一致；
        if tyh.file_exists(self.source_prefix, "卷宗封面"):
            data0 = tyh.file_exists_open(self.source_prefix, "卷宗封面", DocxData)
            result = data0.tabels_content["处理结果"]
            #print(result)
            # 处理结果
            if len(result) == 0:
                table_father.display(self,"处理结果：“处理结果”不应为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "处理结果", "”处理结果“不应为空！")
            else:
                if tyh.file_exists(self.source_prefix, "行政处罚决定书"):
                    #table_father.display(self, "× 请主观审核《卷宗封面》的处理结果和《行政处罚决定书》比对！！", "red")
                    #tyh.addRemarkInDoc(self.mw, self.doc, "处理结果", "请主观审核《行政处罚决定书》以比对！")
                    # 2022.5.23新解决方案
                    result = self.contract_tables_content['处理结果']
                    data = tyh.file_exists_open(self.source_prefix, "行政处罚决定书", DocxData)
                    compare_content = data.text
                    # 先检查是不是原文复制过来的
                    flag = True
                    if result in compare_content:
                        flag = True  # 一致
                    else:
                        # 再提取金额，看金额是否在行政处罚决定书里出现
                        temp = re.findall(r"(\s*\d+\s*.\s*\d+\s*元)", result)
                        if temp:
                            for i in temp:
                                if i in compare_content:
                                    continue
                                else:
                                    flag = False
                                    break
                        else:
                            flag = False
                    if not flag:
                        table_father.display(self, "处理结果：此文书中的“处理结果”和《行政处罚决定书》中“处理结果”不一致，请人工核查！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "处理结果", "此文书中的“处理结果”和《行政处罚决定书》中“处理结果”不一致，请人工核查！")
                else:
                    table_father.display(self, "文书缺失：《行政处罚决定书》.docx不存在！", "red")

            # 共计几页
            #print(data0.text.strip())
            text = data0.text.strip()
            temp = re.findall("此卷共计(.*?)页", text)
            #print(temp[0])
            sp = Simple_Content()
            if sp.is_null(temp[0]):
                table_father.display(self, "共计页数：共计页数 不应为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "此卷共计", "共计页数 不应为空！")
            else:
                page_Num = temp[0].strip()
                #print(page_Num)
                if not page_Num.isdigit():
                    table_father.display(self,"共计页数：“共计页数”应为数字，文中格式不合适！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "此卷共计", "“共计页数”不应为空！")

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
        self.file_name_real = file_name_real
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
        ioc = Table_52(my_prefix, my_prefix)
        contract_file_path = my_prefix + "卷宗封面_.docx"
        ioc.check(contract_file_path, "卷宗封面_.docx")