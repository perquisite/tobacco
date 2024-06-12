#-*- coding:utf-8 -*-
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.TimeOperator import TimeOper
from yancaoRegularDemo.Resource.tools.simple_content import Simple_Content
from yancaoRegularDemo.Resource.tools.EntityRecognition import EntityRecognition


# 结案报告表
from yancaoRegularDemo.Resource.tools.tanweijia_function import is_exist_cover

function_description_dict = {
    'check_Reason': '“案由”：不为空、“ 未在当地烟草专卖批发企业进货、销售非法生产的烟草专卖品、无烟草专卖零售许可证经营烟草制品零售业务、无烟草专卖品准运证运输烟草专卖品、走私烟草专卖品、销售假冒注册商标且伪劣卷烟”，应与案卷中《行政处罚决定书》、《案件处理审批表》、《案件集体讨论记录》、《陈述申辩记录》中“案由一致”。',
    'check_Date': '“立案日期”：不为空，与《立案报告表》中“负责人意见”的落款时间一致，《涉案烟草专卖品核价表》、《调查总结报告》、《延长案件调查终结审批表》、《案件处理审批表》、《卷宗封面》中“立案日期”一致',
    'check_Party': '“当事人”不为空，且与《卷宗封面》《立案报告表》《证据先行登记保存通知书》《抽样取证物品清单》《涉案烟草专卖品核价表》《调查总结报告》《延长案件调查终结审批表》《案件处理审批表》《先行登记保存证据处理通知书》《行政处罚事先告知书》《听证告知书》《当场行政处罚决定书》《行政处罚决定书》《违法物品销毁记录表》《结案报告》中的“当事人”一致',
    'check_CaseSummary': '“案情摘要”：不为空，包含时间、地点、违法事实、调查情况',
    'check_Decision': '“处理决定”：不为空，与《行政处罚决定书》中的处罚决定一致',
    'check_Execution': '“执行情况”：不为空，是否执行完毕',
    'check_UndertakerReason': '“承办人结案理由”：不为空，当事人是否已经接受处罚，是否建议结案',
    'check_UndertakerOpinion': '“承办部门意见”：不为空，是否同意承办人意见，是否建议结案，并注明日期',
    'check_DirectorOpinion': '“负责人意见”：不为空，是否同意结案；负责人签字，并注明日期',
}

class Table51(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prefix = source_prefix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.check_Reason()",
            "self.check_Date()",
            "self.check_Party()",
            "self.check_CaseSummary()",
            "self.check_Decision()",
            "self.check_Execution()",
            "self.check_UndertakerReason()",
            "self.check_UndertakerOpinion()",
            "self.check_DirectorOpinion()"
        ]

    def check_Reason(self):
        file_name = "结案报告表"
        if not tyh.file_exists(self.source_prefix, "结案报告表"):
            table_father.display(self, "文书缺失：《" + file_name + "》不存在", "red")
        else:
            reason0 = self.contract_tables_content["案由"]
            sp = Simple_Content()
            if sp.is_null(reason0):
                table_father.display(self, "案由：案由一栏为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "案    由", "案由一栏不应为空！")
            else:
                reason_list = ["未在当地烟草专卖批发企业进货","销售非法生产的烟草专卖品","无烟草专卖零售许可证经营烟草制品零售业务","无烟草专卖品准运证运输烟草专卖品","走私烟草专卖品、销售假冒注册商标且伪劣卷烟"]
                # 若有多个案由，将其分割
                reason0_list = re.split('[,，、 ]+', reason0)
                for it in reason0_list:
                    if not it in reason_list:
                        table_father.display(self, "案由：案由填写不合法，填写内容应在“未在当地烟草专卖批发企业进货、销售非法生产的烟草专卖品、无烟草专卖零售许可证经营烟草制品零售业务、无烟草专卖品准运证运输烟草专卖品、走私烟草专卖品、销售假冒注册商标且伪劣卷烟”中。", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, it, "案由填写不合法，填写内容应在“未在当地烟草专卖批发企业进货、销售非法生产的烟草专卖品、无烟草专卖零售许可证经营烟草制品零售业务、无烟草专卖品准运证运输烟草专卖品、走私烟草专卖品、销售假冒注册商标且伪劣卷烟”中。")
                # 对比与《行政处罚决定书》的案由
                if not tyh.file_exists(self.source_prefix, "行政处罚决定书"):
                    table_father.display(self, "文书缺失：" + "《行政处罚决定书》.docx不存在", "red")
                else:
                    file_name = "行政处罚决定书_"
                    final_file = ""
                    for root, dirs, files in os.walk(self.source_prefix):
                        for f in files:
                            if file_name + '.docx' == f or file_name + '.doc' == f:
                                final_file = f
                            elif file_name in f and "当场行政处罚决定书_" not in f:
                                final_file = f
                    data = DocxData(self.source_prefix + final_file)
                    #data = tyh.file_exists_open(self.source_prefix, "行政处罚决定书", DocxData)
                    temp = re.search(r"案由([：:])(\s*)(\S*)", data.text)
                    if temp:
                        reason1 = temp.group(3).strip()
                        #print(reason1)
                        if not reason1 == reason0:  # 若reason1含多个案由，则可以改进——分别提取并挨个比较
                            table_father.display(self, "案由：" + "案由与《行政处罚决定书》中的案由（"+reason1+"）不一致！", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, reason0, "《行政处罚决定书》中的案由（"+reason1+"）与此处不一致！")
                    else:
                        table_father.display(self, "案由：没有在《行政处罚决定书》中提取到案由！请人工核查对比", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "", "没有在《行政处罚决定书》中提取到案由！请人工核查对比")

                # 对比与《案件处理审批表》的案由
                if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
                    table_father.display(self, "文书缺失：" + "《案件处理审批表》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
                    reason2 = data.tabels_content['案由']
                    #print(reason2)
                    if not reason2 == reason0:  # 若reason1含多个案由，则可以改进——分别提取并挨个比较
                        table_father.display(self, "案由：" + "案由与《案件处理审批表》中的案由（"+reason2+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, reason0, "《案件处理审批表》中的案由（"+reason2+"）与此处不一致！")
                # 对比与《案件集体讨论记录》的案由
                if not tyh.file_exists(self.source_prefix, "案件集体讨论记录"):
                    table_father.display(self, "文书缺失：" + "《案件集体讨论记录》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "案件集体讨论记录", DocxData)
                    temp = re.search(r"案由([：:])(\s*)(\S*)", data.text)
                    if temp:
                        reason3 = temp.group(3).strip()
                        # print(reason3)
                        if not reason3 == reason0:  # 若reason1含多个案由，则可以改进——分别提取并挨个比较
                            table_father.display(self, "案由：" + "案由与《案件集体讨论记录》中的案由（"+reason3+"）不一致！", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, reason0, "《案件集体讨论记录》中的案由（"+reason3+"）与此处不一致！")
                    else:
                        table_father.display(self, "案由：没有在《案件集体讨论记录》中提取到案由！请人工核查对比", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "", "没有在《案件集体讨论记录》中提取到案由！请人工核查对比")

                # 对比与《陈述申辩记录》的案由
                if not tyh.file_exists(self.source_prefix, "陈述申辩记录"):
                    table_father.display(self, "文书缺失：" + "《陈述申辩记录》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "陈述申辩记录", DocxData)
                    temp = re.search(r"案由([：:])(\s*)(\S*)", data.text)
                    if temp:
                        reason4 = temp.group(3).strip()
                        print(reason4)
                        if not reason4 == reason0:  # 若reason1含多个案由，则可以改进——分别提取并挨个比较
                            table_father.display(self, "案由：" + "案由与《陈述申辩记录》中的案由（"+reason4+"）不一致！", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, reason0, "《陈述申辩记录》中的案由（"+reason4+"）与此处不一致！")
                    else:
                        table_father.display(self, "案由：没有在《陈述申辩记录》中提取到案由！请人工核查对比", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "", "没有在《陈述申辩记录》中提取到案由！请人工核查对比")

                # 对比与《延期（分期）缴纳罚款审批表》的案由    该表暂时么有demo
                # if not tyh.file_exists(self.source_prefix, "延期（分期）缴纳罚款审批表_"):
                #     table_father.display(self, "× " + "《延期（分期）缴纳罚款审批表》.docx不存在", "red")
                # else:

                # 对比与《卷宗封面》的案由
                if not tyh.file_exists(self.source_prefix, "卷宗封面"):
                    table_father.display(self, "文书缺失：" + "《卷宗封面》.docx不存在", "red")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "卷宗封面", DocxData)
                    reason5 = data.tabels_content['案由'].strip()
                    if not reason5 == reason0:  # 若reason1含多个案由，则可以改进——分别提取并挨个比较
                        table_father.display(self, "案由：" + "案由与《卷宗封面》中的案由（"+reason5+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, reason0, "《卷宗封面》中的案由（"+reason5+"）与此处不一致！")

    def check_Date(self):
        date0 = self.contract_tables_content['立案日期']
        sp = Simple_Content()
        t = TimeOper()
        if sp.is_null(date0):
            table_father.display(self, "立案日期：“立案日期”一栏为空！", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "立案日期", "“立案日期”一栏不应为空！")
        else:
            # 与《立案报告表》中“负责人意见”的落款时间一致
            if not tyh.file_exists(self.source_prefix, "立案报告表"):
                table_father.display(self, "文书缺失：" + "《立案报告表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "立案报告表", DocxData)
                temp = re.search(r"日期([：:])(\s*)(\S*)", data.tabels_content['负责人意见'])
                if temp:
                    date1 = temp.group(3).strip()
                    #print(date1)
                    if t.time_order(date0, date1)==0:
                        pass
                    else:
                        table_father.display(self, "立案日期：”立案日期“《立案报告表》中“负责人意见”的落款时间（"+date1+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, date0, "”立案日期“与《立案报告表》中“负责人意见”的落款时间（"+date1+"）不一致！")
                else:
                    table_father.display(self, "立案日期：没有在《立案报告表》中提取到负责人意见的落款时间！请人工核查对比", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "", "没有在《立案报告表》中提取到负责人意见的落款时间！请人工核查对比")

            # 与《涉案烟草专卖品核价表》中“立案日期”一致
            if not tyh.file_exists(self.source_prefix, "涉案烟草专卖品核价表"):
                table_father.display(self, "文书缺失：" + "《涉案烟草专卖品核价表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "涉案烟草专卖品核价表", DocxData)
                temp = re.findall(r"(\S{4})年(\S+)月(\S+)日", data.text)
                #print(temp[-1])
                raw_date = temp[-1][0] + '年' + temp[-1][1] + '月' + temp[-1][2] + '日'
                #print(raw_date)
                date2 = tyh.changeDate(raw_date)
                #print(date2)
                if t.time_order(date0, date2) == 0:
                    pass
                else:
                    table_father.display(self, "立案日期”立案日期“与《涉案烟草专卖品核价表》中的“立案日期”（"+raw_date+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, date0, "”立案日期“与《涉案烟草专卖品核价表》中的“立案日期”（"+raw_date+"）不一致！")

            # 与《调查终结报告》中“立案日期”一致
            if not tyh.file_exists(self.source_prefix, "调查终结报告"):
                table_father.display(self, "文书缺失：" + "《调查终结报告》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "调查终结报告", DocxData)
                date3 = data.tabels_content['立案日期']
                if t.time_order(date0, date3) == 0:
                    pass
                else:
                    table_father.display(self, "立案日期：”立案日期“与《调查终结报告》中的“立案日期”（"+date3+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, date0, "”立案日期“与《调查终结报告》中的“立案日期”（"+date3+"）不一致！")

            # 与《案件处理审批表》中“立案日期”一致
            if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
                table_father.display(self, "文书缺失：" + "《案件处理审批表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
                date4 = data.tabels_content['立案日期']
                if t.time_order(date0, date4) == 0:
                    pass
                else:
                    table_father.display(self, "立案日期：”立案日期“与《案件处理审批表》中的“立案日期”（"+date4+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, date0, "”立案日期“与《案件处理审批表》中的“立案日期”（"+date4+"）不一致！")

            # 与《延长案件调查终结审批表》中“立案日期”一致
            if not tyh.file_exists(self.source_prefix, "延长案件调查终结审批表"):
                table_father.display(self, "文书缺失：" + "《延长案件调查终结审批表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "延长案件调查终结审批表", DocxData)
                date5 = data.tabels_content['立案日期']
                if t.time_order(date0, date5) == 0:
                    pass
                else:
                    table_father.display(self, "立案日期：”立案日期“与《延长案件调查终结审批表》中的“立案日期”（"+date5+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, date0, "”立案日期“与《延长案件调查终结审批表》中的“立案日期”（"+date5+"）不一致！")

            # 与《卷宗封面》中“立案日期”一致
            if not tyh.file_exists(self.source_prefix, "卷宗封面"):
                table_father.display(self, "文书缺失：" + "《卷宗封面》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "卷宗封面", DocxData)
                date6 = data.tabels_content['立案日期']
                if t.time_order(date0, date6) == 0:
                    pass
                else:
                    table_father.display(self, "立案日期：”立案日期“与《卷宗封面》中的“立案日期”（"+date6+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, date0, "”立案日期“与《卷宗封面》中的“立案日期”（"+date6+"）不一致！")

    # 检查当事人
    def check_Party(self):
        party0 = self.contract_tables_content['当事人']
        #print(party0)
        sp = Simple_Content()
        if sp.is_null(party0):
            table_father.display(self, "当事人：当事人一栏为空！", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "当 事 人", "当事人一栏不应为空！")
        else:
            # 与《卷宗封面》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "卷宗封面"):
                table_father.display(self, "文书缺失：" + "《卷宗封面》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "卷宗封面", DocxData)
                party1 = data.tabels_content['当事人']
                #print(party1)
                if party1 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：“当事人”与《卷宗封面》中的“当事人”（"+party1+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《卷宗封面》中的“当事人”（"+party1+"）不一致！")

            # 与《立案报告表》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "立案报告表"):
                table_father.display(self, "文书缺失：" + "《立案报告表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "立案报告表", DocxData)
                party2 = data.tabels_content['当事人']
                #print(party2)
                if party2 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：“当事人”与《立案报告表》中的“当事人”（"+party2+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《立案报告表》中的“当事人”（"+party2+"）不一致！")

            # 与《延长立案期限审批表》中“当事人”一致   该表没有demo

            # 与《证据先行登记保存通知书》中“当事人”一致   先行登记保存证据处理通知书  用到命名实体识别
            if not tyh.file_exists(self.source_prefix, "先行登记保存证据处理通知书"):
                table_father.display(self, "文书缺失：" + "《先行登记保存证据处理通知书》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "先行登记保存证据处理通知书", DocxData)
                raw_data = data.text.strip().split()
                #print(raw_data)
                er = EntityRecognition()
                party4 = ""
                for i in raw_data:
                    if not i == "":
                        temp = er.get_identity_with_tag(i, 'PER')
                        if not temp == []:
                            party4 = temp[0]
                            break
                    else:
                        continue
                if party4 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：“当事人”与《先行登记保存证据处理通知书》中的“当事人”（"+party4+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《先行登记保存证据处理通知书》中的“当事人”（"+party4+"）不一致！")

            # 与《抽样取证物品清单》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "抽样取证物品清单"):
                table_father.display(self, "文书缺失：" + "《抽样取证物品清单》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "抽样取证物品清单", DocxData)
                party5 = data.tabels_content['当事人']
                #print(party5)
                if party5 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：“当事人”与《抽样取证物品清单》中的“当事人”（"+party5+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《抽样取证物品清单》中的“当事人”（"+party5+"）不一致！")

            # 与《涉案烟草专卖品核价表》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "涉案烟草专卖品核价表"):
                table_father.display(self, "文书缺失：" + "《涉案烟草专卖品核价表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "涉案烟草专卖品核价表", DocxData)
                temp = re.search(r"当事人([：:])(\s*)(\S*)(\s*)", data.text)
                party6 = temp.group(3).strip()
                #print(party6)
                if party6 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：“当事人”与《涉案烟草专卖品核价表》中的“当事人”（"+party6+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《涉案烟草专卖品核价表》中的“当事人”（"+party6+"）不一致！")

            # 与《卷烟鉴别检验样品留样、损耗费用审批表》中“当事人”一致   ???此表现无
            if not tyh.file_exists(self.source_prefix, "卷烟鉴别检验样品留样、损耗费用审批表"):
                table_father.display(self, "文书缺失：" + "《卷烟鉴别检验样品留样、损耗费用审批表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "卷烟鉴别检验样品留样、损耗费用审批表", DocxData)
                party7 = data.tabels_content['案件当事人']
                #print(party7)
                if party7 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：“当事人”与《卷烟鉴别检验样品留样、损耗费用审批表》中的“当事人”（"+party7+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《卷烟鉴别检验样品留样、损耗费用审批表》中的“当事人”（"+party7+"）不一致！")

            # 与《调查终结报告》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "调查终结报告"):
                table_father.display(self, "文书缺失：" + "《调查终结报告》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "调查终结报告", DocxData)
                party8 = data.tabels_content['当事人']
                #print(party8)
                if party8 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：“当事人”与《调查终结报告》中的“当事人”（"+party7+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《调查终结报告》中的“当事人”（"+party7+"）不一致！")

            # 与《延长案件调查终结审批表》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "延长案件调查终结审批表"):
                table_father.display(self, "文书缺失：" + "《延长案件调查终结审批表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "延长案件调查终结审批表", DocxData)
                party9 = data.tabels_content['当事人']
                # print(party9)
                if party9 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：“当事人”与《延长案件调查终结审批表》中的“当事人”（"+party9+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《延长案件调查终结审批表》中的“当事人”（"+party9+"）不一致！")

            # 与《案件处理审批表》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
                table_father.display(self, "文书缺失：" + "《案件处理审批表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
                party10 = data.tabels_content['当事人']
                if len(party10) == 2:
                    party10 = party10[1]
                else:
                    pass
                #print(party10)
                if party10 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：“当事人”与《案件处理审批表》中的“当事人”（"+party10+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《案件处理审批表》中的“当事人”（"+party10+"）不一致！")

            # 与《先行登记保存证据处理通知书》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "先行登记保存证据处理通知书"):
                table_father.display(self, "文书缺失：" + "《先行登记保存证据处理通知书》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "先行登记保存证据处理通知书", DocxData)
                raw_data = data.text.strip().split()
                # print(raw_data)
                er = EntityRecognition()
                party11 = ""
                for i in raw_data:
                    if not i == "":
                        temp = er.get_identity_with_tag(i, 'PER')
                        if not temp == []:
                            party11 = temp[0]
                    else:
                        continue
                if party11 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：“当事人”与《先行登记保存证据处理通知书》中的“当事人”（"+party11+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《先行登记保存证据处理通知书》中的“当事人”（"+party11+"）不一致！")

            # 与《行政处罚事先告知书》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "行政处罚事先告知书"):
                table_father.display(self, "文书缺失：" + "《行政处罚事先告知书》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "行政处罚事先告知书", DocxData)
                raw_data = data.text.strip().split()
                # print(raw_data)
                er = EntityRecognition()
                party12 = ""
                for i in raw_data:
                    if not i == "":
                        temp = er.get_identity_with_tag(i, 'PER')
                        if not temp == []:
                            party12 = temp[0]
                    else:
                        continue
                #print(party12)
                if party12 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：当事人与《行政处罚事先告知书》中的“当事人”（"+party12+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "当事人与《行政处罚事先告知书》中的“当事人”（"+party12+"）不一致！")

            # 与《听证告知书》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "听证告知书"):
                table_father.display(self, "文书缺失：" + "《听证告知书》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "听证告知书", DocxData)
                raw_data = data.text.strip().split()
                # print(raw_data)
                er = EntityRecognition()
                party13 = ""
                for i in raw_data:
                    if not i == "":
                        temp = er.get_identity_with_tag(i, 'PER')
                        if not temp == []:
                            party13 = temp[0]
                    else:
                        continue
                #print(party13)
                if party13 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：当事人与《听证告知书》中的“当事人”（"+party13+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "当事人与《听证告知书》中的“当事人”（"+party13+"）不一致！")

            # 与《当场行政处罚决定书》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "当场行政处罚决定书"):
                table_father.display(self, "文书缺失：" + "《当场行政处罚决定书》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "当场行政处罚决定书", DocxData)
                temp = re.search(r"当事人名称（姓名）([：:])(\s*)(\S*)(\s*)", data.text)
                if temp:
                    party14 = temp.group(3).strip()
                    #print(party14)
                    if party14 == party0:
                        pass
                    else:
                        table_father.display(self, "当事人：当事人与《当场行政处罚决定书》中的“当事人”（"+party14+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, party0, "当事人与《当场行政处罚决定书》中的“当事人”（"+party14+"）不一致！")
                else:
                    table_father.display(self, "当事人：未提取到《当场行政处罚决定书》中的“当事人”！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "未提取到《当场行政处罚决定书》中的“当事人”，故无法比较！")

            # 与《行政处罚决定书》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "行政处罚决定书"):
                table_father.display(self, "文书缺失：" + "《行政处罚决定书》.docx不存在", "red")
            else:
                # 防止file_exists_open函数读成《当场行政处罚决定书》，专门写的代码过滤
                # file_name = "行政处罚决定书"
                # for root, dirs, files in os.walk(self.source_prefix):
                #     for f in files:
                #         if file_name + '.docx' == f or file_name + '.doc' == f:
                #             final_file = f
                #         elif file_name in f and "当场行政处罚决定书" not in f:
                #             final_file = f
                file_name = is_exist_cover("行政处罚决定书", "当场行政处罚决定书", self.source_prefix)
                if not file_name:
                    table_father.display(self, "文书缺失：《行政处罚决定书》.docx不存在", "red")
                else:
                    data = DocxData(self.source_prefix + file_name)
                    temp = re.search(r"当事人([：:])(\s*)(\S*)(\s*)", data.text)
                    if temp:
                        party15 = temp.group(3).strip()
                        if party15 == party0:
                            pass
                        else:
                            table_father.display(self, "当事人：当事人与《行政处罚决定书》中的“当事人”（"+party15+"）不一致！", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, party0, "当事人与《行政处罚决定书》中的“当事人”（"+party15+"）不一致！")
                    else:
                        table_father.display(self, "当事人：未提取到《行政处罚决定书》中的“当事人”！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, party0, "未提取到《行政处罚决定书》中的“当事人”，故无法比较！")

            # 与《违法物品销毁记录表》中“当事人”一致
            if not tyh.file_exists(self.source_prefix, "违法物品销毁记录表"):
                table_father.display(self, "文书缺失：" + "《违法物品销毁记录表》.docx不存在", "red")
            else:
                data = tyh.file_exists_open(self.source_prefix, "违法物品销毁记录表", DocxData)
                party16 = data.tabels_content['当事人']
                #print(party16)
                if party16 == party0:
                    pass
                else:
                    table_father.display(self, "当事人：当事人与《违法物品销毁记录表》中的“当事人”（"+party16+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, party0, "当事人与《违法物品销毁记录表》中的“当事人”（"+party16+"）不一致！")

            # 与《加处罚款决定书》中“当事人”一致    无表

            # 与《延期（分期）缴纳罚款审批表》中“当事人”一致    无表
    # 检查 案情摘要
    def check_CaseSummary(self):
        if not tyh.file_exists(self.source_prefix, "结案报告表"):
            pass
            # table_father.display(self, "× " + "《结案报告表》.docx不存在", "red")
        else:
            data = tyh.file_exists_open(self.source_prefix, "结案报告表", DocxData)
            #print(data.tabels_content)
            summary = data.tabels_content['案情摘要'][0]
            #print(summary)
            s = Simple_Content()
            if s.is_null(summary) == "":
                table_father.display(self, "案情摘要：“案情摘要”为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "案情摘要", "案情摘要不应为空！")
            else:
                # 检查包含时间 地点
                er = EntityRecognition()
                temp = er.get_identity_with_tag(summary, 'TIME')
                #print("时间信息\n")
                #print(temp)
                if temp:
                    pass
                else:
                    table_father.display(self, "案情摘要：“案情摘要”并未包含时间信息！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "案情摘要", "“案情摘要”应包含时间信息！")
                temp = er.get_identity_with_tag(summary, 'LOC')
                #print("地点信息\n")
                #print(temp)
                if temp:
                    pass
                else:
                    table_father.display(self, "案情摘要：“案情摘要”并未包含地点信息！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "案情摘要", "“案情摘要”应包含地点信息！")
                # 香烟条数
                if not tyh.file_exists(self.source_prefix, "立案报告表"):
                    table_father.display(self, "文书缺失：" + "《立案报告表》不存在，无法具体对比其“案情摘要”部分与本文书的“案情摘要”部分。", "red")
                else:
                    temp1 = re.findall("\s*(\d+)\s*条", summary)
                    data = tyh.file_exists_open(self.source_prefix, "立案报告表", DocxData)
                    compare = data.tabels_content['案情摘要']
                    temp2 = re.findall("\s*(\d+)\s*条", compare)
                    if temp1 and temp2:
                        #print(temp1)
                        #print(temp2)
                        if len(temp1) == len(temp1):
                            flag = True
                            for i in temp1:
                                if i not in temp2:
                                    flag = False
                                    break
                            if not flag:
                                table_father.display(self,
                                                     "案情摘要：“案情摘要”中的香烟条数条数分别是"+str(temp1)+"，而《立案报告表》的“案情摘要”中香烟条数条数分别是"+str(temp2)+"！",
                                                     "red")
                                tyh.addRemarkInDoc(self.mw, self.doc, "案情摘要",
                                                   "“案情摘要”中的香烟条数条数分别是"+str(temp1)+"，而《立案报告表》的“案情摘要”中香烟条数条数分别是"+str(temp2)+"！")
                        else:
                            table_father.display(self, "案情摘要：“案情摘要”中的香烟条数条数分别是"+str(temp1)+"，而《立案报告表》的“案情摘要”中香烟条数条数分别是"+str(temp2)+"！",
                                                     "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, "案情摘要", "× “案情摘要”中的香烟条数条数分别是"+str(temp1)+"，而《立案报告表》的“案情摘要”中香烟条数条数分别是"+str(temp2)+"！",
                                                     "red")

    # 检查 处理决定
    def check_Decision(self):
        if not tyh.file_exists(self.source_prefix, "结案报告表"):
            pass
            # table_father.display(self, "× " + "《结案报告表》.docx不存在", "red")
        else:
            decision0 = self.contract_tables_content['处理决定'][0]
            # file_name = "行政处罚决定书_"
            # for root, dirs, files in os.walk(self.source_prefix):
            #     for f in files:
            #         if file_name + '.docx' == f or file_name + '.doc' == f:
            #             final_file = f
            #         elif file_name in f and "当场行政处罚决定书_" not in f:
            #             final_file = f
            if tyh.file_exists(self.source_prefix, "行政处罚决定书"):
                data = tyh.file_exists_open(self.source_prefix, "行政处罚决定书", DocxData)
                temp = data.text.split("\n")
                decision1 = ""
                for i in temp:
                    decision1 += i
                # print("decision0\n")
                # print(decision0)
                # print("decision1\n")
                # print(decision1)
                if decision0 in decision1:
                    pass
                else:
                    table_father.display(self, "处理决定：”处理决定“与《行政处罚决定书》中的“处罚决定”（"+decision1+"）不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "处理决定", "“处理决定”与《行政处罚决定书》中的“处罚决定”（"+decision1+"）不一致！")

    # 检查 执行情况
    def check_Execution(self):
        if not tyh.file_exists(self.source_prefix, "结案报告表"):
            pass
            # table_father.display(self, "× " + "《结案报告表》.docx不存在", "red")
        else:
            execution = self.contract_tables_content['执行情况']
            # print(execution)
            if execution == "":
                table_father.display(self, "执行情况：执行情况不应为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "执行情况", "处理决定不应为空！")
            else:
                if "执行了" in execution:
                    pass
                else:
                    table_father.display(self, "执行情况：“执行情况”为“未执行”，请人工核验！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "执行情况", "“执行情况”为“未执行”，请人工核验！")

    # 检查承办人结案理由
    def check_UndertakerReason(self):
        if not tyh.file_exists(self.source_prefix, "结案报告表"):
            pass
            # table_father.display(self, "× " + "《结案报告表》.docx不存在", "red")
        else:
            # print(self.contract_tables_content['承办人结案理由'])
            reason = self.contract_tables_content['承办人结案理由']
            temp = ""
            if isinstance(reason, list):
                for i in reason:
                    temp += i
                reason = temp
            else:
                pass
            reason = reason.replace(" ", "").replace("\n", "").replace(":", "").replace("：", "").replace('\u3000', '').replace("\n", "")
            f = re.search(r"(\S+)签名", reason)
            if not f:
                table_father.display(self, "承办人结案理由：承办人结案理由为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '结案理由', "承办人结案理由不应为空！")
            else:
                # 检查当事人是否已经接受处罚
                temp = re.search(r"已(\S+)处罚", reason)
                #print(temp)
                if temp:
                    pass
                else:
                    table_father.display(self, "承办人结案理由：“承办人结案理由”中当事人未接受处罚，请人工核验！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '结案理', "“承办人结案理由”中当事人未接受处罚，请人工核验！")
                # 检查是否建议结案
                if "建议结案" in reason:
                    pass
                else:
                    table_father.display(self, "是否结案：该案可能未结案，请人工核验！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '结案理由', "该案可能未结案，请人工核验！")
                # 检查 两个承办人签字，并注明日期   （电子签名，跳过）
                t = TimeOper()
                temp = re.search(r"日期(\s*\S+\s*年\s*\S+\s*月\s*\S+\s*日\s*)", reason)
                if temp and t.is_valid_date(temp.group(1)):
                    pass
                else:
                    table_father.display(self, "承办人结案理由：“承办人结案理由”中没有注明日期或格式日期格式不规范！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '结案理由', "”承办人结案理由“中没有注明日期或格式日期格式不规范！")

    # 检查承办部门意见
    def check_UndertakerOpinion(self):
        if not tyh.file_exists(self.source_prefix, "结案报告表"):
            pass
            # table_father.display(self, "× " + "《结案报告表》.docx不存在", "red")
        else:
            #print(self.contract_tables_content)
            opinion = self.contract_tables_content['承办部门意见']
            # print(opinion)
            temp = ""
            if isinstance(opinion, list):
                for i in opinion:
                    temp += i
                opinion = temp
            else:
                pass
            opinion = opinion.replace(" ", "").replace("\n", "").replace(":", "").replace("：", "").replace('\u3000', '').replace("\n", "")
            # print("承办部门意见\n")
            # print(opinion)
            f = re.search(r"(\S+)签名", opinion)
            if not f:
                table_father.display(self, "承办部门意见：”承办部门意见“为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "承办部门", "“承办部门意见”不应为空！")
            else:
                # 是否同意
                if "同意" in opinion:
                    pass
                else:
                    table_father.display(self, "是否同意结案：“承办部门意见”中未审查到同意承办人意见，请人工核验！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "承办部门", "“承办部门意见”中未审查到同意承办人意见，请人工核验！")
                # 是否建议结案
                if "建议结案" in opinion:
                    pass
                else:
                    table_father.display(self, "是否建议结案：“承办部门意见”中未审查到建议结案，请人工核验！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "承办部门", "“承办部门意见”中未审查到建议结案，请人工核验！")
                # 注明日期
                t = TimeOper()
                temp = re.search(r"日期(\s*\S+\s*年\s*\S+\s*月\s*\S+\s*日\s*)", opinion)
                # print("日期\n")
                # print(temp)
                if temp and t.is_valid_date(temp.group(1)):
                    pass
                else:
                    table_father.display(self, "承办部门意见日期：“承办部门意见”中没有注明日期或格式日期格式不规范！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '承办部门', "“承办部门意见”中没有注明日期或格式日期格式不规范！")
    # 检查 负责人意见
    def check_DirectorOpinion(self):
        if not tyh.file_exists(self.source_prefix, "结案报告表"):
            pass
            # table_father.display(self, "× " + "《结案报告表》.docx不存在", "red")
        else:
            opinion = self.contract_tables_content['负责人意见']
            temp = ""
            if isinstance(opinion, list):
                for i in opinion:
                    temp += i
                opinion = temp
            else:
                pass
            opinion = opinion.replace(" ", "").replace("\n", "").replace(":", "").replace("：", "").replace('\u3000', '')
            f = re.search(r"(\S+)签名", opinion)
            if not f:
                table_father.display(self, "负责人意见：负责人意见为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "负责人", "负责人意见不应为空！")
            else:
                # 是否同意
                if "同意" in opinion:
                    pass
                else:
                    table_father.display(self, "负责人意见：“负责人意见”中未审查到同意结案，请人工核验！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "负责人", "“负责人意见”中未审查到同意结案，请人工核验！")
                # 注明日期
                t = TimeOper()
                temp = re.search(r"日期[:：](\s*(\S+)\s*年\s*(\S+)\s*月\s*(\S+)\s*日\s*)", opinion)
                if temp and t.is_valid_date(temp.group(1)):
                    pass
                else:
                    table_father.display(self, "负责人意见：“负责人意见”中没有注明日期或格式日期格式不规范！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '负责人', "“负责人意见”中没有注明日期或格式日期格式不规范！")

    def check(self, contract_file_path, file_name_real):
        print("正在审查"+ file_name_real +"，审查结果如下：")
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
        print("《结案报告表》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = r"C:\\Users\\twj\Desktop\\test1\\"
    lis = os.listdir(my_prefix)
    if "结案报告表_.docx" in lis:
        ioc = Table51(my_prefix, my_prefix)
        contract_file_path = my_prefix + "结案报告表_.docx"
        ioc.check(contract_file_path, "结案报告表_.docx")