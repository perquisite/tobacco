import re
from time import sleep

from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh

function_description_dict = {
    'check_difference': '当事人名称（姓名）、法定代表人（负责人）、性别、民族、证件类型及号码、联系电话、身份证住址、经营地址、烟草专卖许可证号码信息均应与《检查（勘验）笔录》保持一致',
    'check_time': '违法行为时间应当与《检查（勘验）笔录》保持一致。',
    'check_template_one': '处罚依据、结果分类与标准模板不一致。模板格式为：XXXX的行为，违反了XXXXX,现依据XXXXX规定，处以第X项行政处罚。',
    'check_template_two': '向XX烟草专卖局申请行政复议,市州范围内均为统一格式。如果雅安为：雅安市烟草专卖局',
    'check_template_three': '向XXX人民法院起诉。各县局均为统一格式。',
    'check_sign': '应当有两名以上执法人员签字，并注明执法证号。',
    'check_not_empty': '日期、印章、当事人签字不为空。',
}


# 当场行政处罚决定书
class Table35(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prifix = source_prifix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.mw.Visible = 0
        self.mw.DisplayAlerts = 0
        sleep(0.5)
        self.doc = self.mw.Documents.Open(self.my_prefix + "当场行政处罚决定书_.docx")
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.check_difference()",
            "self.check_time()",
            "self.check_template_one()",
            "self.check_template_two()",
            "self.check_template_three()",
            "self.check_sign()",
            "self.check_not_empty()"
        ]

    def check_difference(self):
        """
        作用：当事人名称（姓名）、法定代表人（负责人）、性别、民族、证件类型及号码、
        联系电话、身份证住址、经营地址、烟草专卖许可证号码信息均应与《检查（勘验）笔录》保持一致
        """
        file_name = '检查（勘验）笔录'
        if os.path.exists(self.source_prifix + file_name + "_.docx") == 0:
            table_father.display(self, "文件缺失：《" + file_name + "》不存在", "red")
        else:
            other_text = DocxData(self.source_prifix + file_name + "_.docx").text

            exist_flag_list = [True] * 9
            incident_name_list = ['当事人名称（姓名）', '法定代表人（负责人）', '性别', '民族', '证件类型及号码',
                                  '联系电话', '身份证住址', '经营地址', '烟草专卖许可证号码']
            self_list = []
            other_list = []

            self_name_parttern = re.compile(r'[(当事人名称（姓名）)(当事人)]：\s*([^\s]*)\s*')
            self_presenter_parttern = re.compile(r'法定代表人（负责人）：\s*([^\s]*)\s*')
            self_sex_parttern = re.compile(r'性别：\s*([^\s]*)\s*')
            self_nation_parttern = re.compile(r'民族：\s*([^\s]*)\s*')
            self_certificate_parttern = re.compile(r'[(证件类型及号码)(身份证号)]：\s*([^\s]*)\s*')
            self_phone_parttern = re.compile(r'联系电话：\s*([^\s]*)\s*')
            self_id_address_parttern = re.compile(r'[(身份证住址)(住址)]：\s*([^\s]*)\s*')
            self_business_address_parttern = re.compile(r'经营地址：\s*([^\s]*)\s*')
            self_licence_parttern = re.compile(r'烟草专卖许可证号码：\s*([^\s]*)\s*')

            self_name = re.findall(self_name_parttern, self.contract_text)
            self_presenter = re.findall(self_presenter_parttern, self.contract_text)
            self_sex = re.findall(self_sex_parttern, self.contract_text)
            self_nation = re.findall(self_nation_parttern, self.contract_text)
            self_certificate = re.findall(self_certificate_parttern, self.contract_text)
            self_phone = re.findall(self_phone_parttern, self.contract_text)
            self_id_address = re.findall(self_id_address_parttern, self.contract_text)
            self_business_address = re.findall(self_business_address_parttern, self.contract_text)
            self_licence = re.findall(self_licence_parttern, self.contract_text)
            self_list.append(self_name[0] if self_name != [] else '/')
            self_list.append(self_presenter[0] if self_presenter != [] else '/')
            self_list.append(self_sex[0] if self_sex != [] else '/')
            self_list.append(self_nation[0] if self_nation != [] else '/')
            self_list.append(self_certificate[0] if self_certificate != [] else '/')
            self_list.append(self_phone[0] if self_phone != [] else '/')
            self_list.append(self_id_address[0] if self_id_address != [] else '/')
            self_list.append(self_business_address[0] if self_business_address != [] else '/')
            self_list.append(self_licence[0] if self_licence != [] else '/')

            other_name_parttern = re.compile(r'[(被检查（勘验）人姓名)]：([^\s]*)\s*')
            other_presenter_parttern = re.compile(r'法定代表人（负责人）：([^\s]*)\s*')
            other_sex_parttern = re.compile(r'性别：\s*([^\s]*)\s*')
            other_nation_parttern = re.compile(r'民族：\s*([^\s]*)\s*')
            other_certificate_parttern = re.compile(r'[(证件类型及号码)(身份证号)]：\s*([^\s]*)\s*')
            other_phone_parttern = re.compile(r'联系电话：\s*([^\s]*)\s*')
            other_id_address_parttern = re.compile(r'[住址：]：\s*([^\s]*)\s*')
            other_business_address_parttern = re.compile(r'经营地址：\s*([^\s]*)\s*')
            other_licence_parttern = re.compile(r'烟草专卖许可证号码：\s*([^\s]*)\s*')

            other_name = re.findall(other_name_parttern, other_text)
            other_presenter = re.findall(other_presenter_parttern, other_text)
            other_sex = re.findall(other_sex_parttern, other_text)
            other_nation = re.findall(other_nation_parttern, other_text)
            other_certificate = re.findall(other_certificate_parttern, other_text)
            other_phone = re.findall(other_phone_parttern, other_text)
            other_id_address = re.findall(other_id_address_parttern, other_text)
            other_business_address = re.findall(other_business_address_parttern, other_text)
            other_licence = re.findall(other_licence_parttern, other_text)
            other_list.append(other_name[0] if other_name != [] else '/')
            other_list.append(other_presenter[0] if other_presenter != [] else '/')
            other_list.append(other_sex[0] if other_sex != [] else '/')
            other_list.append(other_nation[0] if other_nation != [] else '/')
            other_list.append(other_certificate[0] if other_certificate != [] else '/')
            other_list.append(other_phone[0] if other_phone != [] else '/')
            other_list.append(other_id_address[0] if other_id_address != [] else '/')
            other_list.append(other_business_address[0] if other_business_address != [] else '/')
            other_list.append(other_licence[0] if other_licence != [] else '/')

            i = 0
            for item in self_list:
                if item == [] or item == [''] or item == '/':
                    exist_flag_list[i] = False
                i += 1
            i = 0
            for item in other_list:
                if item == [] or item == [''] or item == '/':
                    exist_flag_list[i] = False
                i += 1
            i = 0
            for flag in exist_flag_list:
                if not flag:
                    table_father.display(self, incident_name_list[i] + '：在本表或《检查（勘验）笔录》中为空', "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, incident_name_list[i],
                                       incident_name_list[i] + '在本表或《检查（勘验）笔录》中为空')
                elif self_list[i] != other_list[i]:
                    table_father.display(self, str(incident_name_list[i]) + '：与《检查（勘验）笔录》中的记录不同', "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, incident_name_list[i],
                                       str(incident_name_list[i]) + '与《检查（勘验）笔录》中的记录不同')
                else:
                    table_father.display(self, str(incident_name_list[i]) + '：正确。与《检查（勘验）笔录》中的记录相同',
                                         'green')
                i += 1

    def check_time(self):
        """
        作用：违法行为时间应当与《检查（勘验）笔录》保持一致。
        """
        if not tyh.file_exists(self.source_prifix, "检查（勘验）笔录"):
            table_father.display(self, "文件缺失：《检查（勘验）笔录》不存在", "red")
        else:
            flag = True
            other_text = tyh.file_exists_open(self.source_prifix, "检查（勘验）笔录", DocxData).text
            time_parttern = re.compile(r'(\d{4}年\d{1,2}月\d{1,2}日(?:\d{1,2}时)?)')
            self_time_list = re.findall(time_parttern, self.contract_text)
            other_time_list = re.findall(time_parttern, other_text)
            if self_time_list == [] or self_time_list == ['']:
                flag = False
                table_father.display(self,
                                     '违法行为时间：未检索到。请检查时间格式是否规范，时间格式应为xxxx年xx月xx日xx时，中间不得有空格',
                                     'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '当场行政处罚决定书',
                                   '未查找到违法行为时间，请检查时间格式是否规范，时间格式应为xxxx年xx月xx日xx时，中间不得有空格')
            if other_time_list == [] or other_time_list == ['']:
                flag = False
                table_father.display(self, '违法行为时间：未在《检查（勘验）笔录》中未查找到，请检查时间格式是否规范', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '当场行政处罚决定书',
                                   '《检查（勘验）笔录》中未查找到违法行为时间，请检查时间格式是否规范')
            if flag:
                if self_time_list[0] == other_time_list[0]:
                    table_father.display(self, '违法行为时间：正确。与《检查（勘验）笔录》保持一致', 'green')
                else:
                    table_father.display(self, '违法行为时间：与《检查（勘验）笔录》中' + str(other_time_list[0]) + '不一致',
                                         'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, self_time_list[0],
                                       '违法行为时间与《检查（勘验）笔录》不一致，《检查（勘验）笔录》违法行为时间为：' +
                                       other_time_list[0])

    def check_template_one(self):
        """
        作用：对处罚依据、结果分类别与标准模板进行比对，不一致即预警。模板格式为：
             XXXX的行为，违反了XXXXX,现依据XXXXX规定，处以第X项行政处罚。
        """
        template_parttern = re.compile(r'.*的行为[,，]?违反了[^,，]*[,，]?现依据.*规定[,，]?处以第.*项行政处罚')
        result = re.findall(template_parttern, self.contract_text)
        if result == [] or result == ['']:
            table_father.display(self,
                                 '匹配失败：标准模板格式为：XXXX的行为,违反了XXXXX,现依据XXXXX规定,处以第X项行政处罚',
                                 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '当场行政处罚决定书',
                               '标准模板格式为：XXXX的行为,违反了XXXXX,现依据XXXXX规定,处以第X项行政处罚')
        else:
            table_father.display(self, '处罚依据、结果：正确。与标准模板一致', 'green')

    def check_template_two(self):
        """
        作用：向XX烟草专卖局申请行政复议,市州范围内均为统一格式。如果雅安为：雅安市烟草专卖局
        """
        parttern = re.compile(r'向\s*([^\s]*)\s*申请行政复议')
        result = re.findall(parttern, self.contract_text)
        if result == [] or result == ['']:
            table_father.display(self, '匹配失败：标准模板格式为：向XX烟草专卖局申请行政复议', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '当场行政处罚决定书', '标准模板格式为：向XX烟草专卖局申请行政复议')
        else:
            table_father.display(self, '匹配成功：正确。向XX烟草专卖局申请行政复议,为统一格式', 'green')

    def check_template_three(self):
        """
        作用：向XXX人民法院起诉。各县局均为统一格式。
        """
        parttern = re.compile(r'向\s*([^\s]*)\s*人民法院起诉')
        result = re.findall(parttern, self.contract_text)
        if result == [] or result == ['']:
            table_father.display(self, '匹配失败：标准模板格式为：向XXX人民法院起诉', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '当场行政处罚决定书', '标准模板格式为：向XXX人民法院起诉')
        else:
            table_father.display(self, '匹配成功：正确。向XXX人民法院起诉,为统一格式', 'green')

    def check_sign(self):
        """
        作用：应当有两名以上执法人员签字，并注明执法证号。
        """
        parttern = re.compile(
            r'执法人员（签名）：\s*([^\s]*)\s*执法证号：\s*([^\s]*)\s*([^\s]*)\s*执法证号：\s*([^\s]*)\s*[(（]')
        result = re.findall(parttern, self.contract_text)
        if result == [] or result == ['']:
            table_father.display(self, '执法人员（签名）和执法证号：未匹配到', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '当场行政处罚决定书', '未匹配到执法人员（签名）和执法证号')
        else:
            flag = True
            for item in result[0]:
                if item == '':
                    table_father.display(self, '签名数量：应当有两名以上执法人员签字，并注明执法证号。', 'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, '执法人员（签名）',
                                       '应当有两名以上执法人员签字，并注明执法证号。')
                    flag = False
                    break
            if flag:
                table_father.display(self, '签名数量：正确。有两名以上执法人员签字，且已注明执法证号', 'green')

    def check_not_empty(self):
        """
        作用：日期、印章、当事人签字不为空。
        """
        name_parttern = re.compile(r'当事人(签名)：\s*([^\s]*)\s*年')
        result = re.findall(name_parttern, self.contract_text)
        if result == [] or result == ['']:
            table_father.display(self, '当事人签名：不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '当事人(签名)：', '当事人签名不能为空')
        else:
            table_father.display(self, '当事人签名：正确。不为空', 'green')
        tyh.addRemarkInDoc(self.mw, self.doc, '当事人(签名)', '日期、印章不能为空，请主观审查')

    def check(self, contract_file_path, file_name_real):
        print("正在审查《当场行政处罚决定书》，审查结果如下：")
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
        self.doc.Save()
        self.doc.Close()

        self.mw.Quit()
        # self.mw.Quit()
        print("《当场行政处罚决定书》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:/Users/Xie/Desktop/out/"
    list = os.listdir(my_prefix)
    if "当场行政处罚决定书_.docx" in list:
        ioc = Table35(my_prefix, my_prefix)
        contract_file_path = my_prefix + "当场行政处罚决定书_.docx"
        ioc.check(contract_file_path, "当场行政处罚决定书_.docx")
