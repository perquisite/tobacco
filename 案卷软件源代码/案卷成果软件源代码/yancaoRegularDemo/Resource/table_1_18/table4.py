from yancaoRegularDemo.Resource.tools.tangyuhao_readfile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *

function_description_dict = {
    'head': '请核对文书编号（如川烟立XX号）',
    'nameRight': '请检查送达文书名称、受送达人，是否为空',
    'time1Right': '请检查立案期限、落款时间、签收日期的时间格式是否正确',
    'elseRight': '请检查送达人、签收人是否正确',
}


class table4(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix

        self.all_to_check = [
            "self.head()",
            "self.nameRight()",
            "self.time1Right()",
            "self.elseRight()"

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
        pattern = r"烟[[](.*?)[]]延告.*"
        list = re.findall(pattern, text)
        pattern1 = r".*第(.*)号.*"
        list1 = re.findall(pattern1, text)
        if list == [] or list1 == [] or list[0].strip() == "" or list1[0].strip() == "":
            table_father.display(self, "表头：表头不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '延告字第', '表头不能为空')

    def nameRight(self):
        pattern = r"(.*)：[\s\S]*你（单位）涉嫌一案"
        text = re.findall(pattern, self.contract_text)
        text = text[0].strip()
        if text.strip() == "":
            table_father.display(self, "名称：名称不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '你（单位）涉嫌', '名称不能为空')
        else:
            text1 = ""
            # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                data1 = tyh.file_exists_open(self.source_prifix, "立案报告表", file_1)
                text1 = data1.tabels_content["当事人"]
            if text1 != text:
                table_father.display(self, "名称：名称与立案报告表_当事人不同", 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, text1, '名称与立案报告表_当事人不同')
            else:
                table_father.display(self, "名称：名称与立案报告表_当事人相同", 'green')

    def time1Right(self):
        pattern = r".*延长该案立案期限至(.*)。.*"
        time = re.findall(pattern, self.contract_text)[0].strip().replace(" ", "")
        # print(time)
        if tyh.get_strtime(time) == False:
            table_father.display(self, "立案期限：正文中时间格式错误", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '立案期限至', '正文中时间格式错误')
        else:
            table_father.display(self, "立案期限：延长该案立案期限至" + str(time), 'green')

        pattern = r".*[\S\s\n](.*年.*月.*日*)"
        timex = re.findall(pattern, self.contract_text)[1]
        time = re.findall(pattern, self.contract_text)[1].replace(" ", "")
        # print(time)
        if tyh.get_strtime(time) == False:
            table_father.display(self, "落款时间：落款时间格式错误", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, timex, '落款时间格式错误')
        else:
            time = tyh.get_strtime(time)
            # print(time)
            # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                data1 = tyh.file_exists_open(self.source_prifix, "立案报告表", file_1)

                text = data1.tabels_content["负责人意见"]
                if text.strip() == "" or text.strip() == "/":
                    table_father.display(self, "落款时间：立案报告表_负责人意见不为空", "red")
                else:
                    sign, date = tyh.sign_date(text)
                    if date == [""] or date == [] or date[0].strip() == "":
                        table_father.display(self, "落款时间：立案报告表_负责人意见未注明日期", "red")
                    else:
                        date = date[0]
                        time0 = tyh.get_strtime(date)
                        if tyh.time_differ(time, time0) > 7 or tyh.time_differ(time, time0) < 0:
                            table_father.display(self,
                                                 "落款时间：延长立案期限告知书_日期（" + str(time) + "）在立案日期（" + str(
                                                     time0) + "）之前，或在立案日期后（" + str(time0) + "）的七日外", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, timex,
                                               "延长立案期限告知书_日期（" + str(time) + "）在立案日期（" + str(
                                                   time0) + "）之前，或在立案日期后（" + str(time0) + "）的七日外")
                        else:
                            table_father.display(self,
                                                 "落款时间：延长立案期限告知书_日期不为空，日期在立案日期之后，并在立案日期后的七日内",
                                                 "green")
        pattern = r".*签收日期：(.*)"
        time_qianshou = re.findall(pattern, self.contract_text)[0].strip().replace(" ", "")
        if tyh.get_strtime(time_qianshou) == False:
            table_father.display(self, "签收日期：签收时间格式错误", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '签收日期', '签收时间格式错误')
        else:
            if tyh.get_strtime(time) == False:
                pass
            else:
                time_qianshou = tyh.get_strtime(time_qianshou)
                if tyh.time_differ(time_qianshou, time) < 0:
                    table_father.display(self, "签收日期：签收日期不能早于落款日期（" + str(time) + "）", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '签收日期', "签收日期不能早于落款日期（" + str(time) + "）")
                else:
                    table_father.display(self, "签收日期：签收日期晚于落款日期", "green")

    def elseRight(self):
        pattern = r".*送达人：(.*)"
        people = re.findall(pattern, self.contract_text)
        if people == [] or people[0].replace(" ", "") == "":
            table_father.display(self, "送达人：送达人不能为空，至少两哥执法人员签字", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '送达人', '送达人不能为空，至少两哥执法人员签字')
        else:
            table_father.display(self, "送达人：人工审查：送达人至少两哥执法人员签字", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '送达人', '人工审查：送达人至少两哥执法人员签字')

        pattern = r".*签收人：(.*)"
        people = re.findall(pattern, self.contract_text)
        if people == [] or people[0].replace(" ", "") == "":
            table_father.display(self, "签收人：签收人不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '签收人', '签收人不能为空')
        table_father.display(self, "签收人：签收人不为空", 'green')


if __name__ == '__main__':
    # ioc = table3()
    # contract_file_path = my_prefix + "延长立案期限审批表_.docx"
    # ioc.check(contract_file_path)
    my_prefix = "C:\\Users\\Zero\\Desktop\\广安岳池售假岳烟立2021第67号\\"
    l = os.listdir(my_prefix)
    if "延长立案期限告知书_.docx" in l:
        ioc = table4(my_prefix, my_prefix)
        contract_file_path = my_prefix + "延长立案期限告知书_.docx"
        ioc.check(contract_file_path)
