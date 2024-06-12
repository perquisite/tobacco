from yancaoRegularDemo.Resource.tools.tangyuhao_readfile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *

function_description_dict = {
    'head': '请核对文书编号（如川烟立XX号）',
    'nameRight': '请检查烟草专卖局名称是否正确',
    'contentRight': '请依据要求，检查正文部分内容',
    'timeRight': '请检查日期格式',
    'clauseRight': '请检查勘验笔录中是否有法律条款',
}


class table5(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix

        self.all_to_check = [
            "self.head()",
            "self.nameRight()",
            "self.contentRight()",
            "self.timeRight()",
            "self.clauseRight()"

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
        pattern = r"烟辖[[](.*?)[]].*"
        list = re.findall(pattern, text)
        pattern1 = r".*第(.*)号.*"
        list1 = re.findall(pattern1, text)
        if list == [] or list1 == [] or list[0].strip() == "" or list1[0].strip() == "":
            table_father.display(self, "表头：表头不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '烟辖', '表头不能为空')

    def nameRight(self):
        text = self.contract_text
        pattern = r"(.*)烟草专卖局：[\n]"
        name = re.findall(pattern, text)
        if name == [] or name[0].strip() == "":
            table_father.display(self, "名称：名称不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, " 烟草专卖局：", '名称不能为空')
        else:
            name = name[0]
            table_father.display(self, "名称：" + str(name) + "烟草专卖局", "green")

    def contentRight(self):
        pattern = r"关于(.*)一案管辖.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "":
            table_father.display(self, "正文：关于_______一案管辖问题  不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '一案管辖问题', ' 关于_______一案管辖问题  不能为空')
        else:
            text = text[0]
            # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                data1 = tyh.file_exists_open(self.source_prifix, "立案报告表", file_1)

                text0 = data1.tabels_content["当事人"]
                text1 = data1.tabels_content["案由"]
                if text0.strip() == "" or text1.strip() == "":
                    table_father.display(self, "正文：立案报告表中当事人或案由获取失败", "red")
                else:
                    if text0 not in text or text1 not in text:
                        table_father.display(self, "正文：内容与立案报告表中“当事人”（" + str(text0) + "）和“案由”（" + str(
                            text1) + "）不一致", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, '一案管辖问题',
                                           "内容与立案报告表中“当事人”（" + str(text0) + "）和“案由”（" + str(
                                               text1) + "）不一致")
                    else:
                        table_father.display(self, "正文：内容不为空且与立案报告表中当事人和案由一致", "green")

        pattern = r".*因(.*)，根据"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "":
            table_father.display(self, "正文：原因不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '因', ' 原因不能为空')
        else:
            text = text[0]
            if "管辖权争议" not in text:
                table_father.display(self, "正文：原因需要包含“管辖权争议”字眼", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, text, '原因需要包含“管辖权争议”字眼')
            else:
                table_father.display(self, "正文：不为空且包含“管辖权争议”字眼", "green")

    def timeRight(self):
        pattern = r".*[\S\s\n](.*年.*月.*日*)"
        time = re.findall(pattern, self.contract_text)[0].replace(" ", "")
        timex = re.findall(pattern, self.contract_text)[0]
        if tyh.get_strtime(time) == False:
            table_father.display(self, "日期：日期格式错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, timex, '日期格式错误')
        else:
            time = tyh.get_strtime(time)
            # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                data1 = tyh.file_exists_open(self.source_prifix, "立案报告表", file_1)
                time0 = data1.tabels_content["案发时间"]
                pattern = r"(.*日).*"
                time0 = re.findall(pattern, time0)
                if time0 == [] or time0 == [] or time0[0].strip() == "":
                    table_father.display(self, "日期：立案报告表_ 案发日期格式错误", "red")
                else:
                    time0 = time0[0].replace(" ", "")
                    time0 = tyh.get_strtime(time0)
                    # print(time0)
                    if tyh.time_differ(time, time0) < 0:
                        table_father.display(self, "日期：日期在案发日期（" + str(time0) + "）之前", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, timex, "日期在案发日期（" + str(time0) + "）之前")
                    else:
                        table_father.display(self, "日期：日期在案发日期之后", "green")

    def clauseRight(self):
        if "涉嫌" not in self.contract_text:
            tyh.addRemarkInDoc(self.mw, self.doc, '现场情况', '检查勘验笔录中判断未检查到含有法律条款，请主观审查')


if __name__ == '__main__':
    # ioc = table3()
    # contract_file_path = my_prefix + "延长立案期限审批表_.docx"
    # ioc.check(contract_file_path)
    my_prefix = "C:\\Users\\Zero\\Desktop\\副本\\"
    l = os.listdir(my_prefix)
    if "指定管辖通知书_.docx" in l:
        ioc = table5(my_prefix, my_prefix)
        contract_file_path = my_prefix + "指定管辖通知书_.docx"
        ioc.check(contract_file_path, "指定管辖通知书_.docx")
