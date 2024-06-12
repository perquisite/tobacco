import string

import win32com.client

from yancaoRegularDemo.Resource.ReadFile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *

function_description_dict = {
    'peopleAndReasonRight': '当事人名字/名称与《检查（勘验）笔录》中被检查人（被勘验人）一致。',
    'formRight': '承办人由2名执法人员签字；承办人签名时间与负责人签名时间一致。',
    'signAndDateRight': '共计、总计一栏不能为空，且与表格中显示的（几个）品种、（几条）数量一致。',
}


class table7(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix

        self.all_to_check = [
            "self.peopleAndReasonRight()",
            "self.formRight()",
            "self.signAndDateRight()"

        ]

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
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
        self.doc.Save()
        self.doc.Close()
        # self.mw.Quit()
        print(file_name_real + "审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result

    def peopleAndReasonRight(self):
        check_name = []
        # if os.path.exists(self.source_prifix + "检查（勘验）笔录_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "检查（勘验）笔录"):
            data = tyh.file_exists_open(self.source_prifix, "检查（勘验）笔录", DocxData)
            # pattern = r'.*被检查（勘验）人名称：(.*)'
            # text = re.findall(pattern, data.text)
            text = ""
            regex_list = [
                ".*被检查（勘验）人姓名：(.*)性别*",
                ".*被检查（勘验）人姓名：(.*)\n"
            ]
            for regex in regex_list:
                t = re.findall(regex, data.text)
                if t:
                    text = t
                    break
            if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
                table_father.display(self, "检查（勘验） 被检查（勘验）人名称为空", "red")
            else:
                check_name.append(text[0].replace(" ", ""))

        pattern = r".*因(.*)涉嫌.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self, "当事人名字：当事人名字/名称不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, "因", '当事人名字/名称不能为空')
        else:
            text_real = text[0]
            if text_real.replace(" ", "") not in check_name:
                table_father.display(self, "当事人名字：“当事人名字/名称”（" + str(
                    text_real) + "）与《检查（勘验）笔录》中“被检查人（被勘验人）（" + str(
                    check_name[0] if check_name != [] else '未查找到') + "）”不一致", 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, "因", "“当事人名字/名称”（" + str(
                    text_real) + "）与《检查（勘验）笔录》中“被检查人（被勘验人）（" + str(
                    check_name[0] if check_name != [] else '未查找到') + "）”不一致")
            else:
                table_father.display(self, '当事人名字：当事人名字/名称与《检查（勘验）笔录》中被检查人（被勘验人）一致',
                                     'green')

        pattern = r'.*涉嫌(.*)行为.*'
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self, "案由：案由不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, "涉嫌", '案由不能为空')
        else:
            table_father.display(self, "案由：案由不为空", 'green')
            # if os.path.exists(self.my_prefix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                data = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)

                anyou = data.tabels_content["案由"]
                if text[0].replace(" ", "") not in anyou:
                    table_father.display(self, "案由：“案由”与立案报告表中“案由”（" + str(anyou) + "）不一致", 'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, "涉嫌",
                                       "“案由”与立案报告表中“案由”（" + str(anyou) + "）不一致")
                else:
                    table_father.display(self, "案由：案由与立案报告表一致", 'green')

    def formRight(self):
        form = self.contract_tables_content
        type_real = 0.0
        sum_real = 0.0
        for f in range(1, len(form) - 2):
            if form[f][0].strip() == "":
                continue
            else:
                type_real += 1.0
                sum_real += float(form[f][2])

            if form[f][3].strip() == "":
                continue
            else:
                type_real += 1.0
                sum_real += float(form[f][5])

        pattern = r'.*共计：(.*)个.*'
        type = re.findall(pattern, form[-2][0])
        if type == [] or type[0].strip() == "" or type[0].replace(" ", "") == '/':
            table_father.display(self, "表格：共计：（品种）不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "共计：", '共计：（品种）不能为空')
        else:
            type = tyh.ch2num(type[0].replace(" ", ""))
            if type == None:
                table_father.display(self, "表格：共计：（品种）数无法识别", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "总计", '共计：（品种）数无法识别')
            elif float(type) != type_real:
                table_father.display(self, "表格：共计：（品种）数错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "共计：", '共计：（品种）数错误')
            else:
                table_father.display(self, "表格：共计：（品种）数正确", "green")

        pattern = r'.*总计：(.*)条[（]数量[）]'
        sum = re.findall(pattern, form[-2][3])
        if sum == [] or sum[0].strip() == "" or sum[0].replace(" ", "") == '/':
            table_father.display(self, "表格：总计：（数量）不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "总计", '总计：（数量）不能为空')
        else:
            sum = tyh.ch2num(sum[0].replace(" ", ""))
            if sum == None:
                table_father.display(self, "表格：总计：（数量）数无法识别", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "总计", '总计：（数量）数无法识别')
            elif float(sum) != sum_real:
                table_father.display(self, "表格：总计：（数量）数错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "总计", '总计：（数量）数错误')
            else:
                table_father.display(self, "表格：总计：（数量）数正确", "green")

    def signAndDateRight(self):
        pattern = r".*承办人（签名）：(.*)年.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/' or text[0].strip().rstrip(
                string.digits) == "":
            table_father.display(self, "承办人（签名）：承办人（签名）不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, "承办人（签名）", '承办人（签名）不能为空')
        else:
            table_father.display(self, "承办人（签名）：承办人（签名）不为空", 'green')
            tyh.addRemarkInDoc(self.mw, self.doc, "承办人（签名）", '主观审查是否两人签名')

        pattern = r".*负责人意见并签名：(.*)年.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/' or text[0].strip().rstrip(
                string.digits) == "":
            table_father.display(self, "负责人意见并签名：负责人意见并签名：不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, "负责人意见并签名：", '负责人意见并签名：不能为空')
        else:
            table_father.display(self, "负责人意见并签名：负责人意见并签名：不为空", 'green')

        pattern = r".*承办人（签名）：(.*)"
        text = re.findall(pattern, self.contract_text)[0]
        text = text.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip()
        text = tyh.subChar(text)
        date1 = tyh.get_strtime_with_(text)

        pattern = r".*负责人意见并签名：(.*)"
        text = re.findall(pattern, self.contract_text)[0]
        text = text.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip()
        text = tyh.subChar(text)
        date2 = tyh.get_strtime_with_(text)

        if date1 == False:
            table_father.display(self, "承办人（签名）：承办人（签名） 日期错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "承办人（签名）", '承办人（签名） 日期错误')
        else:
            table_father.display(self, "承办人（签名）：承办人（签名） 日期正确", "green")

        if date2 == False:
            table_father.display(self, "负责人意见并签名： 日期错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "负责人意见并签名：", '负责人意见并签名： 日期错误')
        else:
            table_father.display(self, "负责人意见并签名： 日期正确", "green")

        if date1 != False and date2 != False:
            if date1 == date2:
                table_father.display(self, "承办人签名时间：承办人签名时间与负责人签名时间一致。", "green")
            else:
                table_father.display(self, "承办人签名时间：承办人签名时间（" + str(date1) + "）与负责人签名时间（" + str(
                    date2) + "）不一致。", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "年",
                                   "承办人签名时间（" + str(date1) + "）与负责人签名时间（" + str(date2) + "）不一致。")


if __name__ == '__main__':

    my_prefix = "C:/Users/Zero/Desktop/烟草文书demo/2021184117_崇烟立2021第1号/"
    list = os.listdir(my_prefix)
    my_prefix = "C:\\Users/Zero/Desktop/烟草文书demo/2021184117_崇烟立2021第1号/"
    if "证据先行登记保存批准书_.docx" in list:
        ioc = table7(my_prefix, my_prefix)
        contract_file_path = os.path.join(my_prefix, "证据先行登记保存批准书_.docx")
        ioc.check(contract_file_path, "证据先行登记保存批准书_.docx")
