import os
import re
import win32com
from win32com.client import Dispatch
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.ReadFile import DocxData
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
from yancaoRegularDemo.Resource.tools.utils import chinese_to_date

function_description_dict = {
    'check_difference': '当事人基本情况：当事人姓名、性别、年龄、地址信息应当与《案件处理审批表》中记载的一致。'
                        '当事人是自然人的，应载明其姓名、性别、年龄、职业和地址。当事人是法人或者其他组织的，应载明其名称和地址，还应载明法定代表人或者负责人的姓名、职务。',
    'check_case_facts': '案件事实：应当包括案发时间、地点、查获卷烟的品种、数量、金额、违法事实等，内容应当与《案件处理审批表》中“案件事实”记载的一致 。',
    'check_template_one': '主要证据：应当与《证据复制（提取）单》中的说明存在对应关系。',
    'check_template_two': '对行为的定性须符合如下模板：”当事人违法了《中华人民共和国烟草专卖法实施条例》第X条规定，系未在当地烟草专卖批发企业进货的违法行为。"',
    'check_template_three': '对行为的处罚：应当与《案件处理审批》中的“处罚依据”和 “承办人”意见一致。',
    'check_template_four': '行政处罚的履行方式和期限。如：自接到本处罚决定书15日内，应到中国XX银行XX分行XX支行XX分理处缴纳罚款。'
                           '（地址：XX市XX县XX街XX号。联系电话：XXXXXX。）逾期不缴纳罚款，每日按罚款数额的3%加处罚款 。',
    'check_template_five': '不服处罚决定救济的途径和期限。标准模板为：'
                           '如不服本行政处罚决定，可以自收到本决定书之日起六十日内向XXX烟草专卖局申请行政复议,也可以自收到本决定书之日起十五日内直接向XX人民法院起诉。',
    'check_template_six': '作出行政处罚决定的烟草专卖局的名称和日期。日期应在《行政处罚事先告知书》时间3日之后，否则应当预警，如有《听证告知书》，同时应当在《听证告知书》日期3日之后，否则预警。',
}


# 行政处罚决定书
class Table36(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prifix = source_prifix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.mw.Visible = 0
        self.mw.DisplayAlerts = 0
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.check_difference()",
            "self.check_case_facts()",
            "self.check_template_one()",
            "self.check_template_two()",
            "self.check_template_three()",
            "self.check_template_four()",
            "self.check_template_five()",
            "self.check_template_six()"
        ]

    # def _is_consistency(self, list_a, str_b):
    #     if list_a == [] or list_a == [''] or str_b == '':
    #         return 0
    #     elif list_a[0] != str_b:
    #         return 1
    #     else:
    #         return 2

    def check_difference(self):
        """
        作用：当事人基本情况：当事人姓名、性别、年龄、地址信息应当与《案件处理审批表》中记载的一致。
            当事人的基本情况。当事人是自然人的，应载明其姓名、性别、年龄、职业和地址；
            当事人是法人或者其他组织的，应载明其名称和地址，还应载明法定代表人或者负责人的姓名、职务。
        """
        if not tyh.file_exists(self.source_prifix, "案件处理审批表"):
            table_father.display(self, "文件缺失：《案件处理审批表》不存在", "red")
        else:
            othet_tabels_content = tyh.file_exists_open(self.source_prifix, "案件处理审批表", DocxData).tabels_content
            incident_name_list = ['当事人', '性别', '住址']
            self_list = []
            other_list = []
            exist_flag_list = [True, True, True]

            self_name_parttern = re.compile(r'[(当事人)]：\s*([^\s]*)\s*')
            self_sex_parttern = re.compile(r'性别：\s*([^\s]*)\s*')
            self_address_parttern = re.compile(r'[(住址)]：\s*([^\s]*)\s*')

            self_name = re.findall(self_name_parttern, self.contract_text)
            self_sex = re.findall(self_sex_parttern, self.contract_text)
            self_address = re.findall(self_address_parttern, self.contract_text)
            self_list.append(self_name)
            self_list.append(self_sex)
            self_list.append(self_address)

            other_name = othet_tabels_content['当事人']
            other_sex = othet_tabels_content['性别']
            other_address = othet_tabels_content['住址']
            other_list.append(other_name)
            other_list.append(other_sex)
            other_list.append(other_address)

            i = 0
            for item in self_list:
                if item == [] or item == ['']:
                    exist_flag_list[i] = False
                i += 1
            i = 0
            for item in other_list:
                if item == '' or item == ' ':
                    exist_flag_list[i] = False
                i += 1
            i = 0
            for flag in exist_flag_list:
                if not flag:
                    table_father.display(self, str(incident_name_list[i]) + '：在本表或《案件处理审批表》中为空', "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, incident_name_list[i],
                                       str(incident_name_list[i]) + '在本表或《案件处理审批表》中为空')
                elif self_list[i][0] != other_list[i][0]:
                    table_father.display(self,
                                         "【" + str(incident_name_list[i]) + '】信息：与《案件处理审批表》中的记录不同',
                                         "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, incident_name_list[i],
                                       "【" + str(incident_name_list[i]) + '】信息与《案件处理审批表》中的记录不同')
                else:
                    table_father.display(self,
                                         '√ ' + str(incident_name_list[i]) + '：正确。与《案件处理审批表》中的记录相同',
                                         'green')
                i += 1

        if "公司" not in self.contract_text:
            tyh.addRemarkInDoc(self.mw, self.doc, '当事人', '当事人是自然人的，应载明其姓名、性别、年龄、职业和地址。')
        else:
            tyh.addRemarkInDoc(self.mw, self.doc, '当事人',
                               '当事人是法人或者其他组织的，应载明其名称和地址，还应载明法定代表人或者负责人的姓名、职务。')

    def check_case_facts(self):
        """
        作用：案件事实：应当包括案发时间、地点、查获卷烟的品种、数量、金额、违法事实等，内容应当与《案件处理审批表》中“案件事实”记载的一致 。
        """
        if not tyh.file_exists(self.source_prifix, "案件处理审批表"):
            table_father.display(self, "文件缺失：《案件处理审批表》不存在", "red")
        else:
            othet_tabels_content = tyh.file_exists_open(self.source_prifix, "案件处理审批表", DocxData).tabels_content
            if othet_tabels_content['案件事实'] not in self.contract_text:
                tyh.addRemarkInDoc(self.mw, self.doc, '行政处罚决定书',
                                   '案件事实的内容与《案件处理审批表》中记载的“案件事实”不一致，二者需完全一致。')

    def check_template_one(self):
        """
        作用：主要证据：应当与《证据复制（提取）单》中的说明存在对应关系。
        """
        tyh.addRemarkInDoc(self.mw, self.doc, '等证据为证',
                           '主要证据应当与《证据复制（提取）单》中的说明存在对应关系,请人工审查')

    def check_template_two(self):
        """
        作用：对行为的定性：如未在当地烟草专卖批发企业进货的定性标准模板为：当事人违法了《中华人民共和国烟草专卖法实施条例》第 条规定，系未在当地烟草专卖批发企业进货的违法行为。
        """
        parttern = re.compile(
            r'当事人违法了《中华人民共和国烟草专卖法实施条例》第(.*)条规定，系未在当地烟草专卖批发企业进货的违法行为')
        result = re.findall(parttern, self.contract_text)
        if result == [] or result == ['']:
            table_father.display(self,
                                 '对行为的定性：未匹配到。标准模板为：当事人违法了《中华人民共和国烟草专卖法实施条例》第X条规定，系未在当地烟草专卖批发企业进货的违法行为',
                                 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '当场行政处罚决定书',
                               '未匹配到对行为的定性。标准模板为：当事人违法了《中华人民共和国烟草专卖法实施条例》第X条规定，系未在当地烟草专卖批发企业进货的违法行为')

    def check_template_three(self):
        """
        作用：对行为的处罚：应当与《案件处理审批》中的“处罚依据”和 “承办人”意见一致。
        """
        if not tyh.file_exists(self.source_prifix, "案件处理审批表"):
            table_father.display(self, "文件缺失：《案件处理审批表》不存在", "red")
        else:
            othet_tabels_content = tyh.file_exists_open(self.source_prifix, "案件处理审批表", DocxData).tabels_content
            if othet_tabels_content['处罚依据'] not in self.contract_text or othet_tabels_content[
                '承办人意见'] not in self.contract_text:
                tyh.addRemarkInDoc(self.mw, self.doc, '作出下列行政处罚',
                                   '对行为的处罚：应当与《案件处理审批》中的“处罚依据”和 “承办人”意见一致。')

    def check_template_four(self):
        """
        作用：行政处罚的履行方式和期限。如：自接到本处罚决定书15日内，应到中国XX银行XX分行XX支行XX分理处缴纳罚款。（地址：XX市XX县XX街XX号。联系电话：XXXXXX。）逾期不缴纳罚款，每日按罚款数额的3%加处罚款 。
        """
        parttern = re.compile(
            r'自接到本处罚决定书15日内，应到中国.*银行.*分行.*支行(.*)分理处缴纳罚款。（地址：.*市.*县.*街.*号。联系电话：.*）逾期不缴纳罚款，每日按罚款数额的3%加处罚款')
        result = re.findall(parttern, self.contract_text)
        if result == [] or result == ['']:
            table_father.display(self,
                                 '行政处罚的履行方式和期限：未匹配到对。标准模板为：自接到本处罚决定书15日内，应到中国XX银行XX分行XX支行XX分理处缴纳罚款。（地址：XX市XX县XX街XX号。联系电话：XXXXXX。）逾期不缴纳罚款，每日按罚款数额的3%加处罚款',
                                 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '当场行政处罚决定书',
                               '未匹配到对行政处罚的履行方式和期限。标准模板为：自接到本处罚决定书15日内，应到中国XX银行XX分行XX支行XX分理处缴纳罚款。（地址：XX市XX县XX街XX号。联系电话：XXXXXX。）逾期不缴纳罚款，每日按罚款数额的3%加处罚款')

    def check_template_five(self):
        """
        作用：不服处罚决定救济的途径和期限。标准模板为：如不服本行政处罚决定，可以自收到本决定书之日起六十日内向XXX烟草专卖局申请行政复议,也可以自收到本决定书之日起十五日内直接向XX人民法院起诉。
        """
        parttern = re.compile(
            r'如不服本行政处罚决定，可以自收到本决定书之日起六十日内向(.*)烟草专卖局申请行政复议[,，；;]+也可以自收到本决定书之日起十五日内直接向(.*)人民法院起诉')
        result = re.findall(parttern, self.contract_text)
        if result == [] or result == ['']:
            table_father.display(self, '【不服处罚决定】的途径和期限：未识别到解决，请主观审查', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '如不服本行政处罚决定',
                               '未识别到解决【不服处罚决定】的途径和期限。标准模板为：如不服本行政处罚决定，可以自收到本决定书之日起六十日内向XXX烟草专卖局申请行政复议,也可以自收到本决定书之日起十五日内直接向XX人民法院起诉')

    def check_template_six(self):
        """
        作用：作出行政处罚决定的烟草专卖局的名称和日期。日期应在《行政处罚事先告知书》时间3日之后，否则应当预警，如有《听证告知书》，同时应当在《听证告知书》日期3日之后，否则预警。
        """
        time_parttern = re.compile(r'.*(二.{3}年.{1,2}月.{1,3}日).*')
        if not tyh.file_exists(self.source_prifix, "行政处罚事先告知"):
            table_father.display(self, "文件缺失：《行政处罚事先告知书》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "行政处罚事先告知", DocxData)
            other_tabels_text = other_info.text
            this_file_time = re.findall(time_parttern, self.contract_text)[-1]
            other_file_time = re.findall(time_parttern, other_tabels_text)[-1]
            if this_file_time == '' or this_file_time is None:
                table_father.display(self, "本文书作出日期：未找到，作出日期请采用中文年月日格式", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '行政处罚决定书',
                                   '未找到本文书作出日期，作出日期请采用中文年月日格式')
            elif other_file_time == '' or other_file_time is None:
                table_father.display(self, "作出日期：未在《行政处罚事先告知书》未找到，作出日期请采用中文年月日格式",
                                     "red")
            else:
                this_file_time_date = chinese_to_date(this_file_time)
                other_file_time_date = chinese_to_date(other_file_time)
                time_differ = tyh.time_differ(this_file_time_date, other_file_time_date)
                if time_differ < 3:
                    table_father.display(self, "作出日期：应在《行政处罚事先告知》作出之日3天后", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, this_file_time,
                                       "作出日期应在《行政处罚事先告知》作出之日3天后，《行政处罚事先告知》作出日期为：" + other_file_time_date)

        if not tyh.file_exists(self.source_prifix, "听证告知书"):
            table_father.display(self, "文件缺失：《听证告知书》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "听证告知书", DocxData)
            other_tabels_text = other_info.text
            this_file_time = re.findall(time_parttern, self.contract_text)[-1]
            other_file_time = re.findall(time_parttern, other_tabels_text)[-1]
            if this_file_time == '' or this_file_time is None:
                table_father.display(self, "本文书作出日期：未找到，作出日期请采用中文年月日格式", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '行政处罚决定书',
                                   '未找到本文书作出日期，作出日期请采用中文年月日格式')
            elif other_file_time == '' or other_file_time is None:
                table_father.display(self, "作出日期：未在《听证告知书》中匹配，作出日期请采用中文年月日格式", "red")
            else:
                this_file_time_date = chinese_to_date(this_file_time)
                other_file_time_date = chinese_to_date(other_file_time)
                time_differ = tyh.time_differ(this_file_time_date, other_file_time_date)
                if time_differ < 3:
                    table_father.display(self, "作出日期：应在《听证告知书》作出之日3天后", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, this_file_time,
                                       "作出日期应在《听证告知书》作出之日3天后，《听证告知书》作出日期为:" + str(
                                           other_file_time_date))

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
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
        print("《行政处罚决定书》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    if "行政处罚决定书_.docx" in list:
        ioc = Table36(my_prefix, my_prefix)
        contract_file_path = my_prefix + "行政处罚决定书_.docx"
        ioc.check(contract_file_path, '行政处罚决定书_.docx')
