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
from yancaoRegularDemo.Resource.tools.utils import is_valid_date

function_description_dict = {
    'check_Party': '当事人：应当与《案件处理审批表》中当事人一致。',
    'check_CaseNo': '案件编号：应当与《立案报告表》确定的编号一致。',
    'check_DestroyDate': '销毁日期：应当在《行政处罚决定书》《送达回证》送达日期之后，或者《送达公告》日期60日之后，或者表述为“待集中销毁”。',
    'check_DestroyLoc': '销毁地点：不为空。',
    'check_DestroyThings': '“品名、规格型号、单位、数量”应当与《行政处罚决定书》表述为“没收销毁”的物品一致。',
    'check_Null': '其他项目不为空。',
}

class Table44(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prefix = source_prefix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.check_Party()",
            "self.check_CaseNo()",
            "self.check_DestroyDate()",
            "self.check_DestroyLoc()",
            "self.check_DestroyThings()",
            "self.check_Null()"
        ]

    def check_Party(self):
        if not tyh.file_exists(self.source_prefix, "违法物品销毁记录表"):
            table_father.display(self, "文书缺失：" + "《违法物品销毁记录表》.docx不存在", "red")
        else:
            if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
                table_father.display(self, "文书缺失：" + "《案件处理审批表》.docx不存在", "red")
            else:
                party0 = self.contract_tables_content['当事人']
                #print(party0)
                sp = Simple_Content()
                if sp.is_null(party0):
                    table_father.display(self, "当事人：“当事人”一栏为空！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "当事人", "“当事人”一栏为空！")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
                    party1 = data.tabels_content['当事人'][1]
                    #print(party1)
                    if sp.is_consistent(party0, party1):
                        pass
                    else:
                        table_father.display(self, "当事人：“当事人”与《案件处理审批表》中当事人（"+party1+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, party0, "“当事人”与《案件处理审批表》中当事人（"+party1+"）不一致！")

    def check_CaseNo(self):
        if not tyh.file_exists(self.source_prefix, "违法物品销毁记录表"):
            pass
            # table_father.display(self, "× " + "《违法物品销毁记录表》.docx不存在", "red")
        else:
            if not tyh.file_exists(self.source_prefix, "立案报告表"):
                table_father.display(self, "文书缺失：" + "《立案报告表》.docx不存在", "red")
            else:
                num0 = self.contract_tables_content['案件编号']
                #print(num0)
                sp = Simple_Content()
                if sp.is_null(num0):
                    table_father.display(self, "案件编号：“案件编号”一栏为空！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "案件编号", "“案件编号”一栏为空！")
                else:
                    data = tyh.file_exists_open(self.source_prefix, "立案报告表", DocxData)
                    temp = data.text.split()
                    for i in temp:
                        if sp.is_null(i):
                            temp.remove(i)
                    #print(temp)
                    num1 = temp[2]
                    if sp.is_consistent(num0, num1):
                        pass
                    else:
                        table_father.display(self, "案件编号：“案件编号”与《立案报告表》的编号（"+num1+"）不一致！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, num0, "“案件编号”与《立案报告表》的编号（"+num1+"）不一致！")

    # 检查 销毁日期
    def check_DestroyDate(self):
        if not tyh.file_exists(self.source_prefix, "违法物品销毁记录表"):
            pass
            # table_father.display(self, "× " + "《违法物品销毁记录表》.docx不存在", "red")
        else:
            if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
                pass
                # table_father.display(self, "× " + "《案件处理审批表》.docx不存在", "red")
            else:
                date0 = self.contract_tables_content['销毁日期']
                #print(date0)
                sp = Simple_Content()
                t = TimeOper()
                if sp.is_null(date0):
                    table_father.display(self, "销毁日期：“销毁日期”一栏为空！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "销毁日期", "“销毁日期”一栏为空！")
                else:
                    flag = True
                    exist_flag_xingzheng = exist_flag_songdagonggao = exist_flag_songdahuizheng = True
                    if "待集中销毁" not in date0:
                        contain_digit = bool(re.search(r'\d', date0))
                        if contain_digit:
                            date0 = date0[0:4] + '-' + date0[4:6] + '-' + date0[6:8]  # 转化成xxxx-xx-xx的形式
                        else:
                            table_father.display(self, "销毁日期：“销毁日期”应为“待集中销毁”或日期！", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, "销毁日期", "“销毁日期”应为“待集中销毁”或日期！")

                    if "待集中销毁" in date0:
                        pass
                        flag = False
                    #elif tyh.file_exists(self.source_prefix, "送达公告_") and flag:
                    #elif flag:
                    if flag:
                        if tyh.file_exists(self.source_prefix, "送达公告"):
                            data = tyh.file_exists_open(self.source_prefix, "送达公告", DocxData)
                            date3 = data.text.split()[-1]
                            # print(date3)
                            date3 = tyh.changeDate(date3)  # 换成阿拉伯形式
                            # print(date3)
                            if t.time_order(date0, date3) > 60:
                                pass
                            else:
                                table_father.display(self, "销毁日期：“销毁日期”应当在《送达公告》日期（"+date3+"）60日之后！", "red")
                                tyh.addRemarkInDoc(self.mw, self.doc, date0, "“销毁日期”应当在《送达公告》日期（"+date3+"）60日之后！")
                            #flag = False
                        else:
                            exist_flag_songdagonggao = False
                    if flag:
                    # elif flag:
                        tag1 = False
                        tag2 = False
                        final_file = is_exist_cover("行政处罚决定书", "当场行政处罚决定书", self.source_prefix)
                        if final_file == False:
                            exist_flag_xingzheng = False
                            #pass
                            #table_father.display(self, "× 《行政处罚决定书》.docx不存在", "red")
                        else:
                            data = DocxData(self.source_prefix + final_file)
                            date1 = data.text.split()[-2]  # 行政处罚决定书时间
                            date1 = tyh.changeDate(date1)  # 换成阿拉伯形式
                            #print(date1)
                            tag1 = True
                        if not tyh.file_exists(self.source_prefix, "送达回证"):
                            exist_flag_songdahuizheng = False
                            #table_father.display(self, "× " + "《送达回证》.docx不存在", "red")
                            #pass
                        else:
                            data = tyh.file_exists_open(self.source_prefix, "送达回证", DocxData)
                            date2 = data.text.split()[-1]  # 送达回证时间
                            date2 = tyh.changeDate(date2)  # 换成阿拉伯形式
                            #print(date2)
                            tag2 = True
                        if tag1 and tag2:
                            if t.time_order(date0, date1) > 0 and t.time_order(date0, date2) > 0:
                                pass
                            else:
                                table_father.display(self, "销毁日期：“销毁日期”应当在《行政处罚决定书》的送达日期（"+date1+"）、《送达回证》（"+date2+"）的送达日期之后！", "red")
                                tyh.addRemarkInDoc(self.mw, self.doc, date0, "“销毁日期”应当在《行政处罚决定书》的送达日期（"+date1+"）、《送达回证》（"+date2+"）的送达日期之后！")
                            #flag = False
                        else:
                            pass
                    if exist_flag_xingzheng == exist_flag_songdahuizheng == exist_flag_songdagonggao == False:
                        table_father.display(self, "销毁日期：《行政处罚决定书》《送达回证》《送达公告》都不存在，无法判断“销毁日期”的合法性！", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, date0, "《行政处罚决定书》《送达回证》《送达公告》都不存在，无法判断“销毁日期”的合法性！")

    # 检查销毁地点
    def check_DestroyLoc(self):
        if not tyh.file_exists(self.source_prefix, "违法物品销毁记录表"):
            pass
            # table_father.display(self, "× " + "《违法物品销毁记录表》.docx不存在", "red")
        else:
            loc = self.contract_tables_content['销毁地点']
            #print(loc)
            sp = Simple_Content()
            if sp.is_null(loc):
                table_father.display(self, "销毁地点：“销毁地点”不应为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '销毁地点', "“销毁地点”不应为空！")

    # 检查销毁地点 “品名、规格型号、单位、数量”应当与《行政处罚决定书》表述为“没收销毁”的物品一致。
    def check_DestroyThings(self):
        if not tyh.file_exists(self.source_prefix, "违法物品销毁记录表"):
            pass
            # table_father.display(self, "× " + "《违法物品销毁记录表》.docx不存在", "red")
        else:
            # 检查行政处罚决定书是否存在
            file_name = is_exist_cover("行政处罚决定书", "当场行政处罚决定书", self.source_prefix)
            if not file_name:
                pass
                # table_father.display(self, "× " + "《行政处罚决定书》.docx不存在", "red")
            else:
                data = DocxData(self.source_prefix + file_name)
                things = self.contract_tables_content['品名']
                #print(things)
                for key in things:
                    if key in data.text:
                        detail_list = things[key]
                        quantity = round(float(detail_list[2]), 1)
                        unit = detail_list[1]
                        search = key + str(quantity) + unit
                        #print(search)
                        temp = re.search(search, data.text)
                        if temp:
                            pass
                        else:
                            mark = "违法销毁物品" + key + " 的品名、规格型号、单位、数量信息与其在《行政处罚决定书》中的不一致！请人工核查。"
                            table_father.display(self, "违法销毁物品：" + mark, "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, key, mark)
                    else:
                        mark = "违法销毁物品 " + key + " 的信息未在《行政处罚决定书》出现！"
                        table_father.display(self, "违法销毁物品：" + mark, "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, key, mark)
                        continue
    # 检查其他项目不为空：销毁理由、承办人、监销人、证明人， 承办部门签字和负责人签字是电子签名，做不了
    def check_Null(self):
        if not tyh.file_exists(self.source_prefix, "违法物品销毁记录表"):
            pass
            # table_father.display(self, "× " + "《违法物品销毁记录表》.docx不存在", "red")
        else:
            s = Simple_Content()
            if s.is_null(self.contract_tables_content['销毁理由']):
                table_father.display(self, "销毁销毁理由：“销毁销毁理由”不应为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '销毁销毁理由', "“销毁地点”不应为空！")
            if s.is_null(self.contract_tables_content['承办人']):
                table_father.display(self, "承办人：“承办人”不应为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办人', "“承办人”不应为空！")
            if s.is_null(self.contract_tables_content['承办部门签字']):
                table_father.display(self, "承办部门签字：“承办部门签字”不应为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门', "“承办部门签字”不应为空！")
            if s.is_null(self.contract_tables_content['监销人']):
                table_father.display(self, "监销人：“监销人”不应为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '监销人', "“监销人”不应为空！")
            if s.is_null(self.contract_tables_content['证明人']):
                table_father.display(self, "证明人：“证明人”不应为空！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '证明人', "“证明人”不应为空！")

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
        print("《违法物品销毁记录表》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\twj\Desktop\\test\\"
    list = os.listdir(my_prefix)
    if "违法物品销毁记录表_.docx" in list:
        ioc = Table44(my_prefix, my_prefix)
        contract_file_path = my_prefix + "违法物品销毁记录表_.docx"
        ioc.check(contract_file_path, "违法物品销毁记录表_.docx")