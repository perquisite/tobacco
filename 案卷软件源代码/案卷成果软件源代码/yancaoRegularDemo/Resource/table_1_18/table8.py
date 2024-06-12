import string

import win32com.client

from yancaoRegularDemo.Resource.ReadFile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *
from warnings import simplefilter

simplefilter(action='ignore', category=FutureWarning)

function_description_dict = {
    'headRight': '请核对文书编号（如川烟立XX号），“XX烟草专卖局”不能为空。',
    'nameRight': '“X烟存通字[]第X号”（即文书编号）不能为空',
    'reasonRight': '“涉嫌XXX的行为”与《证据先行登记保存批准书》中“涉嫌XX行为”一致。',
    'formRight': '品种规格、数量、共计、总计与《证据先行登记保存批准书》一致。',
    'signAndDateRight': '检查人签字由2名执法人员签字。当事人签字一般与被登记保存人名字/名称一致，若不一致，比如现场负责人代签，出现预警提示。',
}


class table8(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix

        self.all_to_check = [
            "self.headRight()",
            "self.nameRight()",
            "self.reasonRight()",
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

    def headRight(self):
        pattern = "(.*)烟草专卖局.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self, "表头名称：表头名称不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "烟草专卖局", '表头名称不能为空')
        else:
            table_father.display(self, "表头名称：表头名称不为空", "green")

        pattern = r"烟存通字[[](.*?)[]].*"
        list0 = re.findall(pattern, self.contract_text)
        pattern1 = r".*第(.*)号.*"
        list1 = re.findall(pattern1, self.contract_text)
        if list0 == [] or list1 == [] or list0[0].strip() == "" or list1[0].strip() == "":
            table_father.display(self, "表头名称：表头不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '烟存通字', '表头不能为空')

    def nameRight(self):

        pattern = r".*号[\s\S\n](.*)：[\s\S\n]*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self, '表头：名字不能为空', "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '：', '表头不能为空')
        else:
            table_father.display(self, "表头：表头不为空", "green")

    def reasonRight(self):
        pattern = r".*涉嫌(.*)的行为*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self, '正文：原因不能为空', "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '的行为', '原因不能为空')
        else:
            table_father.display(self, "正文：原因不为空", "green")
            # if os.path.exists(self.my_prefix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                data = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
                anyou = data.tabels_content["案由"]
                if text[0].replace(" ", "") not in anyou:
                    table_father.display(self, "正文：“案由”（" + str(
                        text[0].replace(" ", "")) + "）与立案报告表中“案由”（" + str(anyou) + "）不一致", 'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, "涉嫌",
                                       "“案由”（" + str(text[0].replace(" ", "")) + "）与立案报告表中“案由”（" + str(
                                           anyou) + "）不一致")
                else:
                    table_father.display(self, "正文：案由与立案报告表一致", 'green')

    def formRight(self):
        form = self.contract_tables_content
        type_real = 0.0
        sum_real = 0.0
        for f in range(1, len(form) - 1):
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
        type = re.findall(pattern, form[-1][0])
        if type == [] or type[0].strip() == "" or type[0].replace(" ", "") == '/':
            table_father.display(self, "表格：共计：（品种）不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "共计：", '共计：（品种）不能为空')
        else:
            type = tyh.ch2num(type[0].replace(" ", ""))
            if type == None:
                table_father.display(self, "表格：共计：（品种）需要中文大写，如“壹、贰……”", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "总计（数量）", "共计：（品种）需要中文大写，如“壹、贰……”")
            elif float(type) != type_real:
                table_father.display(self, "表格：共计：（品种）数错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "共计：", '共计：（品种）数错误')
            else:
                table_father.display(self, "表格：共计：（品种）数正确", "green")

        pattern = r'.*总计[（]数量[）]：(.*)条.*'
        sum = re.findall(pattern, form[-1][3])
        if sum == [] or sum[0].strip() == "" or sum[0].replace(" ", "") == '/':
            table_father.display(self, "表格：总计：（数量）不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "总计（数量）", '总计：（数量）不能为空')
        else:
            sum = tyh.ch2num(sum[0].replace(" ", ""))
            if sum == None:
                table_father.display(self, "表格：总计：（数量）数无法识别", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "总计（数量）", '总计：（数量）数无法识别')
            elif float(sum) != sum_real:
                table_father.display(self, "表格：总计：（数量）数错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "总计（数量）", '总计：（数量）数错误')
            else:
                table_father.display(self, "表格：总计：（数量）数正确", "green")

        # if os.path.exists(self.source_prifix + "证据先行登记保存通知书_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "证据先行登记保存通知书"):
            data = tyh.file_exists_open(self.source_prifix, "证据先行登记保存通知书", DocxData)
            form0 = data.tabels_content
            # print(form0[:-1])
            # print(form[:-1])
            if form0[:-1] != form[:-1]:
                table_father.display(self,
                                     "品种规格：品种规格、数量、共计、总计与《证据先行登记保存批准书》中的对应信息（" + str(
                                         form0) + "）没有一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "品种规格",
                                   "品种规格、数量、共计、总计与《证据先行登记保存批准书》中的对应信息（" + str(
                                       form0) + "）没有一致")
            else:
                table_father.display(self, "品种规格：品种规格、数量、共计、总计与《证据先行登记保存批准书》一致", "green")

        pattern = "(.*)烟草专卖局.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self, "备注：表头名称不能为空", "red")
        else:
            if text[0] + "烟草专卖局" not in form[-1][0]:
                table_father.display(self, "备注：没有写明证据放在哪里或烟草局与表头不同", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "备注", '没有写明证据放在哪里或烟草局与表头不同')

    def signAndDateRight(self):
        pattern = r".*当事人（签名）：(.*)年.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/' or text[0].strip().rstrip(
                string.digits) == "":
            table_father.display(self, "当事人（签名）：当事人（签名）不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, "当事人（签名）", '当事人（签名）不能为空')
        else:
            table_father.display(self, "当事人（签名）：当事人（签名）不为空", 'green')
            tyh.addRemarkInDoc(self.mw, self.doc, "当事人（签名）", '主观审查是否与被登记保存人名字/名称一致')

        pattern = r".*见证人（签名）：(.*)年.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/' or text[0].strip().rstrip(
                string.digits) == "":
            table_father.display(self, "见证人（签名）：不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, "见证人（签名）：", '见证人（签名）：不能为空')
        else:
            table_father.display(self, "见证人（签名）：不为空", 'green')

        pattern = ".*当事人（签名）：(.*)\n*见证人*"
        text = re.findall(pattern, self.contract_text)[0]
        text = text.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip()
        text = tyh.subChar(text)
        date1 = tyh.get_strtime_with_(text)

        pattern = r".*见证人（签名）：(.*)"
        text = re.findall(pattern, self.contract_text)[0]
        text = text.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip()
        text = tyh.subChar(text)
        date2 = tyh.get_strtime_with_(text)

        if date1 == False:
            table_father.display(self, "当事人（签名）：日期错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "当事人（签名）：", '当事人（签名）： 日期错误')
        else:
            table_father.display(self, "当事人（签名）：日期正确", "green")

        if date2 == False:
            table_father.display(self, "见证人（签名）： 日期错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "见证人（签名）：", '见证人（签名）： 日期错误')
        else:
            table_father.display(self, "见证人（签名）：日期正确", "green")

        if date1 != False and date2 != False:
            if date1 == date2:
                table_father.display(self, "签名时间：承办人签名时间与负责人签名时间一致。", "green")
            else:
                table_father.display(self, "签名时间：承办人签名时间与负责人签名时间不一致。", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "年", '承办人签名时间与负责人签名时间不一致')

        tyh.addRemarkInDoc(self.mw, self.doc, "承办人", '主观审查是否由2名执法人员签字')

        pattern = r".*[\S\s\n](.*年.*月.*日*)"
        timex = re.findall(pattern, self.contract_text)[-1]
        time = re.findall(pattern, self.contract_text)[-1].replace(" ", "")

        if tyh.get_strtime(time) == False:
            table_father.display(self, "落款时间：落款时间格式错误", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, timex, '落款时间格式错误')
        else:
            # if os.path.exists(self.source_prifix + "先行登记保存批准书_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "先行登记保存批准书"):
                data = tyh.file_exists_open(self.source_prifix, "先行登记保存批准书", DocxData)
                text = data.text
                pattern = r".*负责人意见并签名：(.*)"
                text = re.findall(pattern, text)[0]
                text = text.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip()
                text = tyh.subChar(text)
                date2 = tyh.get_strtime_with_(text)
                if date2 == False:
                    table_father.display(self, "落款时间：证据先行登记保存批准书时间格式错误", 'red')
                else:
                    if tyh.get_strtime(time) != date2:
                        table_father.display(self, "落款时间：“年月日”部分（" + str(
                            time) + "）与《证据先行登记保存批准书》作出日期（" + str(date2) + "）不一致", 'red')
                        tyh.addRemarkInDoc(self.mw, self.doc, timex,
                                           "“年月日”部分（" + str(time) + "）与《证据先行登记保存批准书》作出日期（" + str(
                                               date2) + "）不一致")
                    else:
                        table_father.display(self, "落款时间：“年月日”部分与《证据先行登记保存批准书》作出日期一致",
                                             'green')


if __name__ == '__main__':
    my_prefix = "C:\\Users\\12259\\Desktop\\副本\\"
    list = os.listdir(my_prefix)
    if "证据先行登记保存通知书_.docx" in list:
        ioc = table8(my_prefix, my_prefix)
        contract_file_path = my_prefix + "证据先行登记保存通知书_.docx"
        ioc.check(contract_file_path)
