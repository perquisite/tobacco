import os
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.EntityRecognition import EntityRecognition

function_description_dict = {
    'check_cause_of_action': '案由不为空，案由一栏加“涉嫌”两字，一般应与《立案报告表》的案由一致',
    'check_put_on_record_date': '立案日期不为空，立案日期一栏与《立案报告表》立案日期（即负责人意见签名时间）一致。',
    'check_inquirer': '调查人一栏应包括《立案报告表》中的调查人，但可以更多，不能更少。',
    'check_identity_card': '证件类型及号码一般应与《立案报告表》中的“证件类型及号码”一致',
    'check_party': '当事人一栏一般应与《立案报告表》中的当事人一致',
    'check_fact': '调查事实一栏依据主观审查',
    'check_nature_of_the_case': '案件性质一般应与《立案报告表》案由一致',
    'check_basis_for_punishment': '处罚依据一栏细分案由进行讨论',
    'check_handling_opinions': '处理意见不为空，办人签名一栏由2名调查人员签名，日期应在30日内，如有《延长调查终结审批表》时间应在《延长调查终结审批表》中“延长调查终结事由及期限”期限内。',
    'check_additional_function_one_ner': ' 调查事实的无规则文本中是否含有烟草专卖零售许可证、卷烟数目、法律条款等要素',
    'check_additional_funtion_two_ner': '调查事实的无规则文本中是否含有烟草专卖零售许可证、卷烟数目、法律条款等要素',
}


# 案件调查终结报告
class Table20(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prifix = source_prifix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.mw.Visible = 0
        self.mw.DisplayAlerts = 0
        self.contract_text = None
        self.contract_tables_content = None
        self.entityrecognition = EntityRecognition()
        self.target_file_name_list = [
            ['立案报告表_'],
            ['检查（勘验）笔录_']
        ]

        self.all_to_check = [
            "self.check_cause_of_action()",
            "self.check_put_on_record_date()",
            "self.check_inquirer()",
            "self.check_identity_card()",
            "self.check_party()",
            "self.check_fact()",
            "self.check_nature_of_the_case()",
            "self.check_basis_for_punishment()",
            "self.check_handling_opinions()",
            "self.check_additional_function_one_ner()",
            "self.check_additional_funtion_two_ner()"
        ]

    def check_cause_of_action(self):
        """
        作用：案由不为空，案由一栏加“涉嫌”两字，一般应与《立案报告表》的案由一致，若经调查改变案由，出现预警提示。
        """
        # 判断是否为空
        text_cause_of_action = self.contract_tables_content["案由"]
        if text_cause_of_action == "":
            table_father.display(self, "案由：案由不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案   由', '案由不能为空')
        else:
            table_father.display(self, "案由：正确。案由不为空", "green")
            # 判断案由一栏是否加“涉嫌”两字
            if "涉嫌" not in text_cause_of_action:
                table_father.display(self, "案由：案由一栏应该加[涉嫌]两字", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案   由', '案由一栏应该加[涉嫌]两字')
            else:
                table_father.display(self, "案由：正确。案由一栏包含[涉嫌]两字", "green")
        target_file_name_index = self.is_target_file_exit(self.source_prifix, self.target_file_name_list[0])
        if target_file_name_index == -1:
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            # 判断是否与《立案报告表》的案由一致
            register_info = DocxData(
                self.source_prifix + self.target_file_name_list[0][target_file_name_index] + ".docx")
            register_tabels_content = register_info.tabels_content
            if text_cause_of_action != register_tabels_content["案由"]:
                table_father.display(self, "案由：与《立案报告表》的案由不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, text_cause_of_action,
                                   '与《立案报告表》的案由不一致,《立案报告表》的案由为：' + register_tabels_content["案由"])
            else:
                table_father.display(self, "案由：正确。案件调查终结报告_案由 与《立案报告表》的案由一致", "green")

    def check_put_on_record_date(self):
        """
        作用：立案日期不为空，立案日期一栏与《立案报告表》立案日期（即负责人意见签名时间）一致。
        """
        # 判断是否为空
        text_put_on_record_date = self.contract_tables_content["立案日期"]
        if text_put_on_record_date == "":
            table_father.display(self, "立案日期：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '立案日期', '立案日期 不能为空')
        else:
            table_father.display(self, "立案日期：正确。不为空", "green")
            target_file_name_index = self.is_target_file_exit(self.source_prifix, self.target_file_name_list[0])
            if target_file_name_index == -1:
                table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
            else:
                # 判断是否与《立案报告表》的立案日期一致
                # 获取《立案报告表》的负责人意见的立案日期
                register_info = DocxData(
                    self.source_prifix + self.target_file_name_list[0][target_file_name_index] + ".docx")
                register_tabels_content = register_info.tabels_content
                date_pattern = re.compile(r'.*日期：(.*)')
                date_by_principal = re.findall(date_pattern, register_tabels_content["负责人意见"].replace(":", "："))
                if date_by_principal == [''] or date_by_principal == []:
                    table_father.display(self, "立案日期：立案报告表_负责人意见_日期 应具体到XX年XX月XX日", "red")
                    return -1
                date_by_principal = tyh.get_strtime(date_by_principal[0])

                # 获取《案件调查终结报告》的立案日期
                text_put_on_record_date = tyh.get_strtime(text_put_on_record_date)
                if text_put_on_record_date == date_by_principal:
                    table_father.display(self, "立案日期：正确。立案日期与《立案报告表》立案日期（即负责人意见签名时间）一致",
                                         "green")
                else:
                    table_father.display(self, "立案日期：立案日期与《立案报告表》立案日期（即负责人意见签名时间）不一致",
                                         "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '立案日期',
                                       '立案日期与《立案报告表》立案日期（即负责人意见签名时间）不一致，《立案报告表》立案日期为：' + date_by_principal)

    def check_inquirer(self):
        """
        作用：调查人一栏应包括《立案报告表》中的调查人，但可以更多，不能更少。
        """
        # if tyh.file_exists(self.source_prifix, "立案报告表"):
        #     data = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
        tyh.addRemarkInDoc(self.mw, self.doc, '调查人',
                           '请主管审查【调查人】一栏，其应包括《立案报告表》中的调查人，但可以更多，不能更少')

    def check_identity_card(self):
        """
        作用：证件类型及号码一般应与《立案报告表》中的“证件类型及号码”一致，若经调查改变的，出现预警提示。
        """
        target_file_name_index = self.is_target_file_exit(self.source_prifix, self.target_file_name_list[0])
        if target_file_name_index == -1:
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            # 获取《立案报告表》的证件类型及号码
            register_info = DocxData(
                self.source_prifix + self.target_file_name_list[0][target_file_name_index] + ".docx")
            register_tabels_content = register_info.tabels_content
            register_identity_card = register_tabels_content["证件类型及号码"].replace(":", "：")

            # 判断是否与《案件调查终结报告》的证件类型及号码一致
            if register_identity_card == self.contract_tables_content["证件类型及号码"].replace(":", "："):
                table_father.display(self, "证件类型及号码：正确。与《立案报告表》的证件类型及号码一致", "green")
            else:
                table_father.display(self, "证件类型及号码：《立案报告表》的证件类型及号码不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, self.contract_tables_content["证件类型及号码"],
                                   '与《立案报告表》的证件类型及号码不一致,《立案报告表》的证件类型及号码为：' + register_identity_card)

    def check_party(self):
        """
        作用：当事人一栏一般应与《立案报告表》中的当事人一致，若经调查改变当事人，出现预警提示。
        """
        target_file_name_index = self.is_target_file_exit(self.source_prifix, self.target_file_name_list[0])
        if target_file_name_index == -1:
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            # 获取《立案报告表》的当事人
            register_info = DocxData(
                self.source_prifix + self.target_file_name_list[0][target_file_name_index] + ".docx")
            register_tabels_content = register_info.tabels_content
            register_party = register_tabels_content["当事人"]

            # 判断是否与《案件调查终结报告》的当事人一致
            if register_party == self.contract_tables_content["当事人"]:
                table_father.display(self, "当事人：正确。与《立案报告表》的当事人一致", "green")
            else:
                table_father.display(self, "当事人：与《立案报告表》的当事人不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, self.contract_tables_content["当事人"],
                                   '与《立案报告表》的当事人不一致，《立案报告表》的当事人为：' + register_party)

    def check_fact(self):
        """
        作用：调查事实一栏依据主观审查
        """
        table_father.display(self, "调查事实：提示。本栏依据主观审查", "green")
        tyh.addRemarkInDoc(self.mw, self.doc, '调查\r事实', '调查事实一栏依据主观审查')

    def check_nature_of_the_case(self):
        """
        作用：案件性质一般应与《立案报告表》案由一致，若经调查改变案由，出现预警提示。
        """
        target_file_name_index = self.is_target_file_exit(self.source_prifix, self.target_file_name_list[0])
        if target_file_name_index == -1:
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            # 获取《立案报告表》的案由
            register_info = DocxData(
                self.source_prifix + self.target_file_name_list[0][target_file_name_index] + ".docx")
            register_tabels_content = register_info.tabels_content
            register_nature_of_the_case = register_tabels_content["案由"]

            # 判断是否与《案件调查终结报告》的案件性质一致
            if register_nature_of_the_case == self.contract_tables_content["案件性质"]:
                table_father.display(self, "案件性质：正确。与《立案报告表》的案由一致", "green")
            else:
                table_father.display(self, "案件性质：与《立案报告表》的案由不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案件\r性质',
                                   '与《立案报告表》的案由不一致,《立案报告表》的案由为：' + register_nature_of_the_case)

    def check_basis_for_punishment(self):
        """
        作用：处罚依据一栏细分案由进行讨论
        """
        table_father.display(self, "处罚依据：提示，本栏细分案由进行讨论", "green")
        tyh.addRemarkInDoc(self.mw, self.doc, '处罚\r依据', '处罚依据一栏细分案由进行讨论')

    def check_handling_opinions(self):
        """
        作用：处理意见不为空，得出是否已经查明事实，建议调查终结，
        承办人签名一栏由2名调查人员签名，日期应在30日内，如有《延长调查终结审批表》时间应在《延长调查终结审批表》中“延长调查终结事由及期限”期限内。
        """
        text_handling_opinions = self.contract_tables_content["处理意见"]
        handling_opinions_pattern = re.compile(r'(.*)签名')
        handling_opinions = re.findall(handling_opinions_pattern, text_handling_opinions)
        if handling_opinions == [''] or handling_opinions == []:
            table_father.display(self, "处理意见：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '处理\r意见', '处理意见不能为空')
        else:
            table_father.display(self, "处理意见：正确。不为空", "green")
            if "查明事实" in handling_opinions[0] or "建议调查终结" in handling_opinions[0]:
                table_father.display(self, "调查事实：正确。已经查明事实", "green")
            else:
                table_father.display(self, "调查事实：未得出是否已经查明事实", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '处理\r意见', '未得出是否已经查明事实')
        tyh.addRemarkInDoc(self.mw, self.doc, '处理\r意见', '请主观检查【承办人签名】一栏，此处应由2名调查人员签名')
        tyh.addRemarkInDoc(self.mw, self.doc, '处理\r意见', '请主观检查【日期】一栏，应在立案日期30日内')

    def check_additional_function_one_ner(self):
        """
        date:2022.2.20
        function:
        1.调查事实中的日期应与检查（勘验）时间的开始时间一致
        2.调查事实中的地点应与检查（勘验）地点一致
        3.调查事实中是否包含执法人员
        """
        target_file_name_index = self.is_target_file_exit(self.source_prifix, self.target_file_name_list[1])
        if target_file_name_index == -1:
            table_father.display(self, "文件缺失：《检查（勘验）笔录》不存在", "red")
        else:
            # 获取《检查（勘验）笔录》的时间与地点部分文字
            other_context = DocxData(
                self.source_prifix + self.target_file_name_list[1][target_file_name_index] + ".docx").text
            other_time = re.findall("检查（勘验）时间：(.*?)\n", other_context)[0]
            other_address = re.findall("检查（勘验）地点：(.*?)\n", other_context)[0]
            # 使用NER获取案件调查终结报告的时间，取第一个出现
            cognitio = self.contract_tables_content['调查事实']
            this_time_for_check = self.entityrecognition.get_identity_with_tag(cognitio, "TIME")[0]
            # print(this_time_for_check)

            if this_time_for_check in other_time:
                table_father.display(self, "调查事实：正确。日期与检查（勘验）时间的开始时间一致", "green")
            else:
                table_father.display(self,
                                     '调查事实：日期为' + this_time_for_check + ',与检查（勘验）时间的开始时间不一致',
                                     "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '调查\r事实',
                                   '调查事实中的日期为' + this_time_for_check + ',与检查（勘验）时间的开始时间不一致')

            if other_address in cognitio:
                table_father.display(self, "调查事实：正确。地点与检查（勘验）地点一致", "green")
            else:
                table_father.display(self, "调查事实：地点与检查（勘验）地点不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '调查\r事实', '调查事实中的地点与检查（勘验）地点不一致')

            if "执法人员" not in cognitio:
                table_father.display(self, "调查事实：不包含执法人员", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '调查\r事实', '调查事实中不包含执法人员')

    def check_additional_funtion_two_ner(self):
        """
        date:2022.2.20
        function:
        调查事实的无规则文本中是否含有烟草专卖零售许可证、卷烟数目、法律条款等要素
        """
        cognitio = self.contract_tables_content['调查事实']
        if "烟草专卖零售许可证" not in cognitio:
            table_father.display(self, "调查事实：不包含烟草专卖零售许可证", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '调查\r事实', '调查事实中不包含烟草专卖零售许可证')
        if "卷烟数目" not in cognitio:
            table_father.display(self, "调查事实：不包含卷烟数目", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '调查\r事实', '调查事实中不包含卷烟数目')
        if "法律条款" not in cognitio:
            table_father.display(self, "调查事实：不包含法律条款", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '调查\r事实', '调查事实中不包含法律条款')

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
        print("《案件调查终结报告》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    dit_list = os.listdir(my_prefix)
    if "案件调查终结报告_.docx" in dit_list:
        ioc = Table20(my_prefix, my_prefix)
        contract_file_path = my_prefix + "案件调查终结报告_.docx"
        ioc.check(contract_file_path, "案件调查终结报告_.docx")
