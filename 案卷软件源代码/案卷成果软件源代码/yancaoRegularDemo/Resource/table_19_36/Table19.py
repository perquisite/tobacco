import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh

function_description_dict = {
    'check_difference': '判断案由、案件来源、案发时间、案发地点、当事人、证件类型及号码、地址、案情摘要是否与《立案报告表》一致',
    'check_comments_of_the_undertaker': '承办人意见不为空，写明撤案的具体缘由：a.经调查认为当事人的违法事实不能成立，不得给予行政处罚的；b'
                                        '.当事人的违法行为情节轻微，依法可以不予行政处罚的；c.案件需要移送给其他机关的。明确是否建议撤销立案，并由两名承办人签字，签署日期，日期在立案日期之后。',
    'check_comments_of_the_department': '承办部门不为空，明确是否同意承办人意见，签名，盖部门印章，时间与“承办人意见”一栏时间一致，或者晚于“承办人意见”一栏时间',
    'check_comments_of_the_principal': '负责人意见不为空，明确是否同意撤销立案，并由负责人签字，盖单位印章，日期与“承办部门意见”一栏的时间一致，或者晚于承办部门意见日期',
}


# 撤销立案报告表
class Table19(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prifix = source_prifix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.mw.Visible = 0
        self.mw.DisplayAlerts = 0
        # self.mw.Visible = 0
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.check_difference()",
            "self.check_comments_of_the_undertaker()",
            "self.check_comments_of_the_department()",
            "self.check_comments_of_the_principal()"
        ]

    def check_difference(self):
        """
        作用：判断案由、案件来源、案发时间、案发地点、当事人、证件类型及号码、地址、案情摘要是否与《立案报告表》一致。
        """
        if not tyh.file_exists(self.source_prifix, "立案报告表"):
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            register_info = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
            register_tabels_content = register_info.tabels_content
            ready_to_check = ["案由", "案件来源", "案发时间", "案发地点", "当事人", "证件类型及号码", "地址",
                              "案情摘要"]
            for i in ready_to_check:
                if register_tabels_content[i] == self.contract_tables_content[i]:
                    table_father.display(self, i + "：" + "正确。" + i + "与《立案报告表》一致", "green")
                else:
                    table_father.display(self, i + "：" + i + "与《立案报告表》不一致", "red")
                    if i == "案由":
                        tyh.addRemarkInDoc(self.mw, self.doc, "案    由",
                                           i + "与《立案报告表》不一致,《立案报告表》的" + i + "为：" +
                                           register_tabels_content[i])
                    elif i == "案发时间":
                        tyh.addRemarkInDoc(self.mw, self.doc, "案  发 时 间",
                                           i + "与《立案报告表》不一致,《立案报告表》的" + i + "为：" +
                                           register_tabels_content[i])
                    elif i == "当事人":
                        tyh.addRemarkInDoc(self.mw, self.doc, "当    事  人",
                                           i + "与《立案报告表》不一致,《立案报告表》的" + i + "为：" +
                                           register_tabels_content[i])
                    else:
                        tyh.addRemarkInDoc(self.mw, self.doc, i, i + "与《立案报告表》不一致,《立案报告表》的" + i + "为：" +
                                           register_tabels_content[i])

    def check_comments_of_the_undertaker(self):
        """
        作用：承办人意见不为空，写明撤案的具体缘由：
        a.经调查认为当事人的违法事实不能成立，不得给予行政处罚的；
        b.当事人的违法行为情节轻微，依法可以不予行政处罚的；
        c.案件需要移送给其他机关的。明确是否建议撤销立案，并由两名承办人签字，签署日期，日期在立案日期之后。
        """
        # 判断承办人意见是否为空
        text_undertaker = self.contract_tables_content["承办人意见"]
        text_pattern = re.compile(r'(.*)签名')
        comments = re.match(text_pattern, text_undertaker)
        if comments.group(1) == "" or comments.group(1) is None:
            table_father.display(self, "承办人意见：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承 办 人\r意 见', '承办人意见不能为空')
        else:
            table_father.display(self, "承办人意见：正确。已填写", "green")
            # 判断是否写明撤案的具体缘由
            ready_to_check = ["违法事实不能成立", "不得给予行政处罚", "违法行为情节轻微", "可以不予行政处罚",
                              "需要移送给其他机关", "建议撤销立案"]
            whether_written = False
            for i in ready_to_check:
                if i in text_undertaker:
                    whether_written = True
                    table_father.display(self, "承办人意见：正确。已写明撤案的具体缘由", "green")
                    if i == "需要移送给其他机关" or i == "建议撤销立案":

                        # 判断签名是否由两名承办人签字
                        first_name_pattern = re.compile(r'.*签名：(.*)[,，、]')
                        second_name_pattern = re.compile(r'签名：.*[,，、](.*)日期.*')
                        first_name = re.findall(first_name_pattern, text_undertaker)
                        second_name = re.findall(second_name_pattern, text_undertaker)
                        if first_name == [''] or first_name == []:
                            table_father.display(self, "承办人意见_承办人签名：需由两名承办人签字", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, '承 办 人\r意 见', '承办人签名需由两名承办人签字')
                        elif second_name == [''] or second_name == []:
                            table_father.display(self, "承办人意见_承办人签名：需由两名承办人签字", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, '承 办 人\r意 见', '承办人签名需由两名承办人签字')
                        else:
                            table_father.display(self, "承办人意见_承办人签名：正确。已由两名承办人签字", "green")

                        # 判断是否签署日期及日期合法性
                        if os.path.exists(self.source_prifix + "立案报告表_.docx") == 0:
                            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
                        else:

                            # 获取立案时间
                            register_info = DocxData(self.source_prifix + "立案报告表_.docx")
                            register_tabels_content = register_info.tabels_content
                            register_date = register_tabels_content["案发时间"]
                            if register_date == [''] or register_date == []:
                                table_father.display(self, "《立案报告表》案发时间：应具体到XX年XX月XX日XX时XX分", "red")
                            register_date = tyh.get_strtime(register_date[0])

                            # 获取撤销立案时间
                            text_undertaker = text_undertaker.replace(":", "：")
                            revocation_date_pattern = re.compile(r'.*日期：(.*)')
                            revocation_date_by_undertaker = re.findall(revocation_date_pattern, text_undertaker)
                            if revocation_date_by_undertaker == [''] or revocation_date_by_undertaker == []:
                                table_father.display(self, "承办人意见_日期：应具体到XX年XX月XX日", "red")
                                tyh.addRemarkInDoc(self.mw, self.doc, '承 办 人\r意 见', '日期应具体到XX年XX月XX日')
                            revocation_date_by_undertaker = tyh.get_strtime(revocation_date_by_undertaker[0])

                            # 判断撤销立案日期是否在立案日期之后
                            if register_date is not False and revocation_date_by_undertaker is not False:
                                date_differ = tyh.time_differ(revocation_date_by_undertaker, register_date)
                                if date_differ > 0:
                                    table_father.display(self, "承办人意见_日期：正确。在立案日期之后", "green")
                                else:
                                    table_father.display(self, "承办人意见_日期：日期不应该在立案日期之前", "red")
                                    tyh.addRemarkInDoc(self.mw, self.doc, '承 办 人\r意 见', '日期不应该在立案日期之前')
                    break
            if whether_written is False:
                table_father.display(self, "承办人意见：未写明撤案的具体缘由", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承 办 人\r意 见', '未写明撤案的具体缘由')

    def check_comments_of_the_department(self):
        """
        作用：承办部门不为空，明确是否同意承办人意见，签名，盖部门印章，时间与“承办人意见”一栏时间一致，或者晚于“承办人意见”一栏时间。
        """
        # 判断承办部门意见是否为空
        text_department = self.contract_tables_content["承办部门意见"]
        text_department = text_department.replace(":", "：")
        text_pattern = re.compile(r'(.*)签名')
        comments = re.match(text_pattern, text_department)
        if comments.group(1) == "" or comments.group(1) is None:
            table_father.display(self, "承办部门意见：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '承办部门意见 不能为空')
        else:
            table_father.display(self, "承办部门意见：正确。不为空", "green")

            # 判断是否同意承办人意见
            if "同意" not in text_department:
                table_father.display(self, "承办部门意见：未明确是否同意承办人意见", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '未明确是否同意承办人意见')
            else:
                table_father.display(self, "承办部门意见：正确。已明确是否同意承办人意见", "green")

            # 判断是否签名
            name_pattern = re.compile(r'.*签名：(.*)日期.*')
            name = re.findall(name_pattern, text_department)
            if name == [''] or name == []:
                table_father.display(self, "承办部门意见：签名不能为空", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '签名不能为空')
            else:
                table_father.display(self, "承办部门意见：正确。签名已填写", "green")
            # 明确是否盖章，进行提示
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '请检查是否加盖部门印章')

            # 判断时间是否与“承办人意见”一栏时间一致，或者晚于“承办人意见”一栏时间
            # 获取承办人意见_时间
            text_undertaker = self.contract_tables_content["承办人意见"].replace(":", "：")
            revocation_date_pattern = re.compile(r'.*日期：(.*)')
            revocation_date_by_undertaker = re.findall(revocation_date_pattern, text_undertaker)
            if revocation_date_by_undertaker == [''] or revocation_date_by_undertaker == []:
                table_father.display(self, "承办人意见_日期：应具体到XX年XX月XX日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承 办 人\r意 见', '日期应具体到XX年XX月XX日')
            revocation_date_by_undertaker = tyh.get_strtime(revocation_date_by_undertaker[0])

            # 获取承办部门意见_时间
            revocation_date_by_department = re.findall(revocation_date_pattern, text_department)
            if revocation_date_by_department == [''] or revocation_date_by_department == []:
                table_father.display(self, "承办部门意见_日期：应具体到XX年XX月XX日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '日期应具体到XX年XX月XX日')
            revocation_date_by_department = tyh.get_strtime(revocation_date_by_department[0])

            # 判断时间顺序
            if revocation_date_by_undertaker is not False and revocation_date_by_department is not False:
                date_differ = tyh.time_differ(revocation_date_by_department, revocation_date_by_undertaker)
                if date_differ >= 0:
                    table_father.display(self, "承办部门意见_日期：正确。在【承办人意见_日期】之后或同一天", "green")
                else:
                    table_father.display(self, "承办部门意见_日期：在【承办人意见_日期】之前", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见',
                                       '承办部门意见_日期 在 承办人意见_日期 之前，承办人意见_日期为：' + revocation_date_by_undertaker)

    def check_comments_of_the_principal(self):
        """
        作用：负责人意见不为空，明确是否同意撤销立案，并由负责人签字，盖单位印章，日期与“承办部门意见”一栏的时间一致，或者晚于承办部门意见日期。
        """
        # 判断负责人意见是否为空
        text_principal = self.contract_tables_content["负责人意见"]
        text_principal = text_principal.replace(":", "：")
        text_pattern = re.compile(r'(.*)签名')
        comments = re.match(text_pattern, text_principal)
        if comments.group(1) == "" or comments.group(1) is None:
            table_father.display(self, "负责人意见：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '负责人意见 不能为空')
        else:
            table_father.display(self, "负责人意见：正确。不为空", "green")

            # 判断是否同意撤销立案
            if "同意" not in text_principal:
                table_father.display(self, "负责人意见：未明确是否同意撤销立案", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '未明确是否同意撤销立案')
            else:
                table_father.display(self, "负责人意见：正确。已明确是否同意撤销立案", "green")

            # 判断是否由负责人签字
            name_pattern = re.compile(r'.*签名：(.*)日期.*')
            name = re.findall(name_pattern, text_principal)
            if name == [''] or name == []:
                table_father.display(self, "负责人意见：签名不能为空", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '签名不能为空')
            else:
                table_father.display(self, "负责人意见：正确。签名已填写", "green")

            # 明确是否盖章，进行提示
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '请检查是否盖单位印章')

            # 判断日期是否与“承办部门意见”一栏的时间一致，或者晚于承办部门意见日期。
            # 获取承办部门意见_时间
            text_department = self.contract_tables_content["承办部门意见"].replace(":", "：")
            revocation_date_pattern = re.compile(r'.*日期：(.*)')
            revocation_date_by_department = re.findall(revocation_date_pattern, text_department)
            if revocation_date_by_department == [''] or revocation_date_by_department == []:
                table_father.display(self, "承办部门意见_日期：应具体到XX年XX月XX日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '日期应具体到XX年XX月XX日')
            revocation_date_by_department = tyh.get_strtime(revocation_date_by_department[0])

            # 获取负责人意见_时间
            revocation_date_by_principal = re.findall(revocation_date_pattern, text_principal)
            if revocation_date_by_principal == [''] or revocation_date_by_principal == []:
                table_father.display(self, "负责人意见_日期：应具体到XX年XX月XX日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '日期应具体到XX年XX月XX日')
            revocation_date_by_principal = tyh.get_strtime(revocation_date_by_principal[0])

            # 判断时间顺序
            if revocation_date_by_principal is not False and revocation_date_by_department is not False:
                date_differ = tyh.time_differ(revocation_date_by_principal, revocation_date_by_department)
                if date_differ >= 0:
                    table_father.display(self, "负责人意见_日期：正确。在【承办部门意见_日期】之后或同一天", "green")
                else:
                    table_father.display(self, "负责人意见_日期：在【承办部门意见_日期】之前", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见',
                                       '负责人意见的日期，不能在承办部门意见的日期之前,承办部门意见的日期为：' + revocation_date_by_department)

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
                table_father.display(self, "文档格式有误，请主观审查下列功能："+function_description_dict[str(func)[5:-2]], "red")
                table_father.display(self, "文档存在格式错误，函数失效：" + func + ' 遇到错误:' + str(e.args))
        self.doc.Save()
        self.doc.Close()

        self.mw.Quit()
        print("《撤销立案报告表》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    if "撤销立案报告表_.docx" in list:
        ioc = Table19(my_prefix, my_prefix)
        contract_file_path = my_prefix + "撤销立案报告表_.docx"
        ioc.check(contract_file_path, '撤销立案报告表_.docx')

