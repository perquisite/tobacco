from yancaoRegularDemo.Resource.tools.tangyuhao_readfile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *

function_description_dict = {
    'head': '请核对文书编号（如川烟立XX号）',
    'reasonRight': '1.应完整填写当事人涉嫌违法行为的性质，如“涉嫌未在当地烟草专卖批发企业进货、涉嫌无烟草专卖品准运证运输烟草专卖品”等;'
                   '2.违法行为前应填写“涉嫌”二字;'
                   '3.违法行为性质表述应与法律条款保持一致;'
                   '4.案由应与卷宗封面、立案报告表、行政处罚决定书、证据先行登记保存通知书、委托鉴定告知书、询问（调查）通知书等文书保持一致',
    'basicInfOfPeopleAndAbstractOfCase': '案发地点、当事人、证件类型及号码、地址、案情摘要应与立案报告表相关内容保持一致',
    'opinionsOfTheDepartment': '1.承办部门负责人应作出是否同意延长或不予延长的意见;2.承办部门负责人应签名并注明日期',
    'opinionsOfPersonInCharge': '1.承办案件的烟草专卖行政主管部门负责人签署审批意见;2. 承办案件的烟草专卖行政主管部门负责人应签名并注明日期',
}


class table3(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix

        self.all_to_check = [
            "self.head()",
            "self.reasonRight()",
            "self.basicInfOfPeopleAndAbstractOfCase()",
            "self.opinionsOfTheDepartment()",
            "self.opinionsOfPersonInCharge()",
        ]

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
        self.mw = win32com.client.Dispatch("Word.Application")
        self.doc = self.mw.Documents.Open(self.my_prefix + file_name_real)
        data = file_1(file_path=contract_file_path)
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
        self.doc.Save()
        self.doc.Close()
        # self.mw.Quit()
        print(file_name_real + "审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result

    def head(self):
        text = self.contract_text
        pattern = r"烟[[](.*?)[]]延立.*"
        list = re.findall(pattern, text)
        pattern1 = r".*第(.*)号.*"
        list1 = re.findall(pattern1, text)
        if list == [] or list1 == [] or list[0].strip() == "" or list1[0].strip() == "":
            table_father.display(self, "表头：表头不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '延长立案期限审批表', '表头不能为空')

    def reasonRight(self):
        reason = self.contract_tables_content["案由"]
        reason0 = ""
        # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "立案报告表"):
            data1 = tyh.file_exists_open(self.source_prifix, "立案报告表", file_1)

            text0 = data1.tabels_content["案由"]
            if text0 != "":
                reason0 = text0
        if reason.strip() == "" or reason.strip() == "/":
            table_father.display(self, "案由：违法行为不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案   由', '违法行为不能为空')
        elif "涉嫌" not in reason:
            table_father.display(self, "案由：违法行为前未填写“涉嫌”二字", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案   由', '违法行为前未填写“涉嫌”二字')
        elif reason != reason0:
            table_father.display(self, "案由：案由与立案报告表等不一致", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案   由', '案由与立案报告表等不一致')
        else:
            table_father.display(self, "案由：案由格式正确", "green")

    def basicInfOfPeopleAndAbstractOfCase(self):
        people = self.contract_tables_content["当事人"]

        people0 = ""
        id0 = ""
        place0 = ""
        abstract0 = ""
        # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "立案报告表"):
            data1 = tyh.file_exists_open(self.source_prifix, "立案报告表", file_1)
            text0 = data1.tabels_content["当事人"]
            if text0 != "":
                people0 = text0
            text0 = data1.tabels_content["证件类型及号码"]
            if text0 != "":
                id0 = text0
            text0 = data1.tabels_content["地址"]
            if text0 != "":
                place0 = text0
            text0 = data1.tabels_content["案情摘要"]
            if text0 != "":
                abstract0 = text0
        if people.strip() == "" or people.strip() == "/":
            table_father.display(self, "当事人：当事人不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '当事人', '当事人不能为空')
        elif people != people0:
            table_father.display(self, "当事人：“当事人”没有与立案报告表中“当事人”（" + str(people0) + "）保持一致", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '当事人',
                               "“当事人”没有与立案报告表中“当事人”（" + str(people0) + "）保持一致")

        id = self.contract_tables_content["证件类型及号码"]
        right_id = id0

        id.replace(' ', '')
        right_id.replace(' ', '')

        if id.strip() == "" or id.strip() == "/":
            table_father.display(self, "当事人：证件类型及号码不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '证件类型及号码', '证件类型及号码不能为空')
        elif id != right_id:
            table_father.display(self, "当事人：“证件类型及号码”没有与立案报告表中证件类型及号码（" + str(
                right_id) + "）保持一致", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '证件类型及号码',
                               "“证件类型及号码”没有与立案报告表中证件类型及号码（" + str(right_id) + "）保持一致")

        place = self.contract_tables_content["地址"]
        right_place = place0

        if place.strip() == "" or place.strip() == "/":
            table_father.display(self, "当事人：地址不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '地  址', '地址不能为空')
        elif place != right_place:
            table_father.display(self, "当事人：“地址”没有与立案报告表中的“地址”（" + str(right_place) + "）保持一致",
                                 "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '地  址',
                               "“地址”没有与立案报告表中的“地址”（" + str(right_place) + "）保持一致")

        # phone = 延长立案期限审批表_联系电话
        # if phone == "":
        #     table_father.display(self,"电话不能为空", "red")
        # else:
        #     s = Simple_Content()
        #     if s.match_re(s.pattern_strings["phone_number"], phone) == []:
        #         table_father.display(self,"电话格式出现错误", "red")
        #     else:
        #         table_father.display(self,"电话格式正确", "green")
        abstract = self.contract_tables_content["案情摘要"]
        right_abstract = abstract0
        if abstract.strip() == "" or abstract.strip() == "/":
            table_father.display(self, "案情摘要：案情摘要不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要', '案情摘要不能为空')
        else:
            abstract = abstract.replace(' ', '')
            right_abstract = right_abstract.replace(' ', '')
            if abstract != right_abstract:
                table_father.display(self, "案情摘要：“案情摘要”没有与与立案报告表中“案情摘要”（" + str(
                    right_abstract) + "）内容保持一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要',
                                   "“案情摘要”没有与与立案报告表中“案情摘要”（" + str(right_abstract) + "）内容保持一致")

        reason = self.contract_tables_content["延长立案期限事由及期限"]
        if isinstance(reason, type(list)): reason = reason[0]
        pattern = r"(.*)签名.*"
        reason0 = re.findall(pattern, reason)
        if reason0 == [] or reason0 == [''] or reason0[0].strip() == "":
            table_father.display(self, "延长立案期限事由及期限：延长立案期限事由及期限不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '延长立案期限事由及期限', '延长立案期限事由及期限不能为空')
        else:
            table_father.display(self, "延长立案期限事由及期限：1.延长立案的理由主观审查 2.延长立案的期限主观审查",
                                 "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '延长立案期限事由及期限',
                               '1.延长立案的理由主观审查 2.延长立案的期限主观审查')

        sign, date = tyh.sign_date(reason)
        if sign == [""] or sign == [] or sign[0].strip() == "":
            table_father.display(self, "延长立案期限事由及期限：延长立案期限事由及期限未签名", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '延长立案期限事由及期限', '延长立案期限事由及期限未签名')
        if date == [""] or date == [] or date[0].strip() == "":
            table_father.display(self, "延长立案期限事由及期限：延长立案期限事由及期限未注明日期", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '延长立案期限事由及期限', '延长立案期限事由及期限未注明日期')

    def opinionsOfTheDepartment(self):
        opinions = self.contract_tables_content["承办部门意见"]
        if isinstance(opinions, list): opinions = opinions[0]
        pattern = r"(.*)签名.*"
        opinions0 = re.findall(pattern, opinions)
        if opinions0 == [] or opinions0 == [""] or opinions0[0].strip() == "":
            table_father.display(self, "承办部门意见：承办部门意见不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '承办部门意见不能为空')
        else:
            if "同意延长" in opinions:
                table_father.display(self, "承办部门意见：同意延长", "green")
            elif "不予延长" in opinions:
                table_father.display(self, "承办部门意见：不予延长", "green")
            else:
                table_father.display(self, "承办部门意见：未明确表明是否延长", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '未明确表明是否延长')

        sign, date = tyh.sign_date(opinions)
        if sign == [''] or sign == [] or sign[0].strip() == "":
            table_father.display(self, "承办部门意见：承办部门意见未签名", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '承办部门意见未签名')
        if date == [""] or date == [] or date[0].strip() == "":
            table_father.display(self, "承办部门意见：承办部门意见未注明日期", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '承办部门意见未注明日期')

    def opinionsOfPersonInCharge(self):
        opinions = self.contract_tables_content["负责人意见"]
        if isinstance(opinions, list): opinions = opinions[0]
        pattern = r"(.*)签名.*"
        opinions0 = re.findall(pattern, opinions)
        if opinions0 == [] or opinions0 == [""] or opinions0[0].strip() == "":
            table_father.display(self, "负责人意见：负责人意见不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '负责人意见不能为空')
        sign, date = tyh.sign_date(opinions)
        if sign == [''] or sign == [] or sign[0].strip() == "":
            table_father.display(self, "负责人意见：负责人意见未签名", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '负责人意见未签名')
        if date == [""] or date == [] or date[0].strip() == "":
            table_father.display(self, "负责人意见：负责人意见未注明日期", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '负责人意见未注明日期')


if __name__ == '__main__':
    # ioc = table3()
    # contract_file_path = my_prefix + "延长立案期限审批表_.docx"
    # ioc.check(contract_file_path)
    my_prefix = "C:\\Users\\12259\\Desktop\\原\\副本_data2\\"
    l = os.listdir(my_prefix)
    if "延长立案期限审批表_.docx" in l:
        ioc = table3(my_prefix, my_prefix)
        contract_file_path = my_prefix + "延长立案期限审批表_.docx"
        ioc.check(contract_file_path)
