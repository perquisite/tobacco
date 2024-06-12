import string
from time import sleep

import win32com.client

from yancaoRegularDemo.Resource.ReadFile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *
from warnings import simplefilter

simplefilter(action='ignore', category=FutureWarning)

function_description_dict = {
    'peopleAndId': '1、当事人：应当与《立案报告表》中的当事人一致；2、立案编号：应当与《立案报告表》中的编号一致。',
    'proposeAndTimeAndPlace': '1、抽样目的：不为空;2、抽样时间：应精确到分，格式为XXXX年XX月XX日XX时XX分。时间规则需要进一步讨论;3、抽样地点：不为空',
    'formElseRight': '1.品种规格：应当与《证据先行登记保存批准书》中的“品种规格”一致;2.样品基数：应当与《证据先行登记保存批准书》中对应的数量一致;'
                     '3.抽样数量：【样品基数在100条-500条的，抽样数量为2条】【样品基数在2条-100条的，抽样数量为1条或2条】'
                     '【样品基数在2条以下的，抽样数量为1条或等于样品基数】【样品基数在500条-2500条的，抽样数量为2-5条】'
                     '【样品基数在2500条以上的，抽样数量为5-10条】',
    'signRight': '1.承办人：应当由两人以上签字;2.负责人意见并签名：应当有“同意”或“不同意”字样并签名;'
                 '3.当事人：签名不为空;4.如当事人签名为空，则见证人签名不能为空',
}


# 9.抽样取证物品清单
class table9(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix

        self.all_to_check = [
            "self.peopleAndId()",
            "self.proposeAndTimeAndPlace()",
            "self.formElseRight()",
            "self.signRight()"

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
        self.mw.Quit()
        print(file_name_real + "审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result

    def peopleAndId(self):
        '''
        1、当事人：应当与《立案报告表》中的当事人一致。
        2、立案编号：应当与《立案报告表》中的编号一致。
        '''
        form = self.contract_tables_content
        people = form["当事人"]
        id = form["立案编号"]

        if people.strip() == "":
            table_father.display(self, "当事人：当事人不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '当事人', '当事人不能为空')

        if id.strip() == "":
            table_father.display(self, "立案编号：立案编号不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '立案编号', '立案编号不能为空')

        if people.strip() != "" and id.strip() != "":
            table_father.display(self, "当事人：当事人不为空", "green")
            table_father.display(self, "立案编号：立案编号不为空", "green")
            # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                data = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
                form0 = data.tabels_content
                text0 = data.text
                people0 = form0["当事人"]
                id0 = ""

                pattern1 = r".*立案报告表[\S\s\n](.*第.*号)[\S\s\n].*"
                list1 = re.findall(pattern1, text0)
                if list1 == [] or list1[0].strip() == "" or list1[0].replace(" ", "") == '/':
                    table_father.display(self, "立案报告表表头为空", "red")
                else:
                    id0 = list1[0]

                if people != people0:
                    table_father.display(self, "当事人：“当事人”（" + str(people) + "）与立案报告表中的“当事人”（" + str(
                        people0) + "）不一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '当事人',
                                       "“当事人”（" + str(people) + "）与立案报告表中的“当事人”（" + str(
                                           people0) + "）不一致")
                else:
                    table_father.display(self, "当事人：当事人与立案报告表一致", "green")

                if id != id0:
                    table_father.display(self, "立案编号：“立案编号”（" + str(id) + "）与立案报告表中的编号（" + str(
                        id0) + "）不一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '立案编号',
                                       "“立案编号”（" + str(id) + "）与立案报告表中的编号（" + str(id0) + "）不一致")
                else:
                    table_father.display(self, "立案编号：立案编号与立案报告表一致", "green")

    def proposeAndTimeAndPlace(self):
        '''
        3、抽样目的：不为空
        4、抽样时间：应精确到分，格式为XXXX年XX月XX日XX时XX分。时间规则需要进一步讨论。
        5、抽样地点：不为空
        '''
        form = self.contract_tables_content
        propose = form["抽样目的"]
        time = form["抽样时间"]
        place = form["抽样地点"]

        if propose.strip() == "":
            table_father.display(self, "抽样目的：抽样目的为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '抽样目的', '抽样目的为空')
        else:
            table_father.display(self, "抽样目的：抽样目的不为空", "green")

        if tyh.get_strtime_5(time) == False:
            table_father.display(self, "抽样时间：抽样时间格式不正确", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '抽样时间', '抽样时间格式不正确')
        else:
            table_father.display(self, "抽样时间：抽样时间格式正确", "green")

        if place.strip() == "":
            table_father.display(self, "抽样地点：抽样地点为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '抽样地点', '抽样地点为空')
        else:
            table_father.display(self, "抽样地点：抽样地点不为空", "green")

    def formElseRight(self):
        '''

        '''
        doc = docx.Document(self.my_prefix + "抽样取证物品清单_.docx")
        all_cell = []
        for t in doc.tables:
            for row in t.rows:
                cells = []
                for cell in row.cells:
                    cells.append(
                        cell.text.replace("\n", "").replace("\t", "").replace("\r", "").replace(" ", ""))
                all_cell.append(cells)
        form = all_cell
        # print(form)

        form = form[4:]
        if form[0][0].strip() == "" or form[0][1].strip() == "" or form[0][2].strip() == "" or form[0][3].strip() == "":
            table_father.display(self, "品种规格：品种规格 或 包装形式 或 样品基数（条） 或 抽样数量（条） 不能为空")
            tyh.addRemarkInDoc(self.mw, self.doc, '品种规格',
                               '品种规格 或 包装形式 或 样品基数（条） 或 抽样数量（条） 不能为空')
            return
        dict = {}
        i = 0
        flag = 1
        form_flag = 1
        while "备注" not in form[i][0]:
            if form[i][0].strip() == '':
                i += 1
                continue
            dict[form[i][0]] = float(form[i][2])
            if form[i][1].strip() not in ["条盒硬盒", "条盒软盒", "硬盒", "软盒", "条包硬盒",
                                          "条包软盒"] and form_flag == 1:
                table_father.display(self, "包装形式：包装形式错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "包装形式", '包装形式错误')
                form_flag = 0
            if self.numRight(form[i][2].strip(), form[i][3].strip()) == 0:
                table_father.display(self, "样品基数：样品基数 与 抽样数量 对应错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "样品基数", '样品基数 与 抽样数量 对应错误')
                flag = 0

            i += 1
        # print(dict)
        if flag:
            table_father.display(self, "样品基数：样品基数 与 抽样数量 对应正确", "green")

        # if os.path.exists(self.source_prifix + "先行登记保存批准书_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "证据先行登记保存批准书"):
            doc = tyh.file_exists_open(self.source_prifix, "证据先行登记保存批准书", docx.Document)
            all_cell = []
            for t in doc.tables:
                for row in t.rows:
                    cells = []
                    for cell in row.cells:
                        cells.append(
                            cell.text.replace("\n", "").replace("\t", "").replace("\r", "").replace(" ", ""))
                    all_cell.append(cells)
            form0 = all_cell

            dict1 = {}
            form0 = form0[1:]
            i = 0
            while "共计" not in form0[i][0]:
                if form0[i][0].strip() != "":
                    dict1[form0[i][0]] = float(form0[i][2])
                if form0[i][3].strip() != "":
                    dict1[form0[i][3]] = float(form0[i][5])
                i += 1
            # print(dict1)

            if dict != dict1:
                table_father.display(self, "品种规格：品种规格或样品基数（条）（" + str(
                    dict) + "）与证据先行登记保存批准书中的对应信息（" + str(dict1) + "）不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '品种规格', "品种规格或样品基数（条）（" + str(
                    dict) + "）与证据先行登记保存批准书中的对应信息（" + str(dict1) + "）不一致")
            else:
                table_father.display(self, "品种规格：品种规格或样品基数（条）与证据先行登记保存批准书一致", "green")

    def numRight(self, num1, num2):
        num1 = float(num1)
        num2 = float(num2)
        flag = 1
        if num1 <= 2 and num2 not in [1.0, 2.0]: flag = 0
        if 2 <= num1 <= 100 and num2 not in [1.0, 2.0]: flag = 0
        if 100 <= num1 <= 500 and num2 != 2.0: flag = 0
        if 500 <= num1 <= 2500 and num2 not in [2.0, 3.0, 4.0, 5.0]: flag = 0
        if 2500 <= num1 and num2 not in [5.0, 6.0, 7.0, 8.0, 9.0, 10.0]: flag = 0
        return flag

    def signRight(self):
        pattern = ".*承办人（签名）：(.*)执法证号*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self, "承办人（签名）：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办人（签名）：', '承办人（签名）：不能为空')
        else:
            table_father.display(self, "承办人（签名）：不为空", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办人（签名）：', '主观审查是否有两个承办人（签名）')

        pattern = ".*执法证号：(.*)"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/' or text[1].strip() == "" or text[
            1].replace(" ", "") == '/':
            table_father.display(self, "执法证号：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '执法证号：', '执法证号：不能为空')
        else:
            table_father.display(self, "执法证号：不为空", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, '执法证号：', '主观审查是否有两个执法证号')

        pattern = r".*负责人意见并签名：(.*)年.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/' or text[0].strip().rstrip(
                string.digits) == "":
            table_father.display(self, "负责人意见并签名：不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, "负责人意见并签名：", '负责人意见并签名：不能为空')
        else:
            table_father.display(self, "负责人意见并签名：不为空", 'green')
            if "同意" not in text and "不同意" not in text:
                table_father.display(self, "负责人意见并签名：没有同意或者不同意字样", 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, "负责人意见并签名：", '负责人意见并签名：没有同意或者不同意字样')

        form = self.contract_tables_content
        time = form["抽样时间"]
        date0 = []
        if tyh.get_strtime_5(time) != False:
            pattern = r"(.*日).*"
            text = re.findall(pattern, time)[0]
            text = text.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip()
            date0 = tyh.subChar(text)
            date0 = date0.split("-")
            for i in range(0, len(date0)):
                date0[i] = int(date0[i])
        else:
            table_father.display(self, "抽样时间：抽样时间错误", 'red')

        pattern = r".*负责人意见并签名：(.*)"
        text = re.findall(pattern, self.contract_text)[0]
        date1 = tyh.get_strtime(text)
        if date1 == False:
            table_father.display(self, "负责人意见：负责人意见并签名 日期错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "负责人意见并签名", '负责人意见并签名 日期错误')
        else:
            out_put_date = date1
            date1 = date1.split("-")
            for i in range(0, len(date1)):
                date1[i] = int(date1[i])
            if date0 != date1:
                table_father.display(self, "负责人意见：负责人意见并签名 日期与抽样时间不同", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "负责人意见并签名",
                                   '负责人意见并签名 日期与抽样时间【' + out_put_date + '】不同')
            else:
                table_father.display(self, "负责人意见：负责人意见并签名 日期正确", "green")

        pattern = r".*当事人（签名）：(.*)年.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/' or text[0].strip().rstrip(
                string.digits) == "":
            table_father.display(self, "当事人（签名）：不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, "当事人（签名）：", '当事人（签名）：不能为空')
        else:
            table_father.display(self, "当事人（签名）：不为空", 'green')

        pattern = r".*当事人（签名）：(.*)"
        text = re.findall(pattern, self.contract_text)[0]
        date1 = tyh.get_strtime(text)
        if date1 == False:
            table_father.display(self, "当事人（签名）： 日期错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "当事人（签名）：", '当事人（签名）： 日期错误')
        else:
            out_put_date = date1
            date1 = date1.split("-")
            for i in range(0, len(date1)):
                date1[i] = int(date1[i])
            if date0 != date1:
                table_father.display(self, "当事人（签名）： 日期与抽样时间不同", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "当事人（签名）：",
                                   '当事人（签名）： 日期与抽样时间【' + out_put_date + '】不同')
            else:
                table_father.display(self, "当事人（签名）： 日期正确", "green")

        pattern = r".*见证人（签名）：(.*)年.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/' or text[0].replace("、",
                                                                                                     "").strip().rstrip(
            string.digits) == "":
            table_father.display(self, "见证人（签名）：不能为空", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, "见证人（签名）：", '见证人（签名）：不能为空')
        else:
            table_father.display(self, "见证人（签名）：不为空", 'green')

        pattern = r".*见证人（签名）：(.*年.*月.*日*)"
        text = re.findall(pattern, self.contract_text)[0]
        date1 = tyh.get_strtime(text)

        if date1 == False:
            table_father.display(self, "见证人（签名）： 日期错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "见证人（签名）：", '见证人（签名）： 日期错误')
        else:
            out_put_date = date1
            date1 = date1.split("-")
            for i in range(0, len(date1)):
                date1[i] = int(date1[i])
            if date0 != date1:
                table_father.display(self, "见证人（签名）：日期与抽样时间不同", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "见证人（签名）：",
                                   '见证人（签名）： 日期与抽样时间【' + out_put_date + '】不同')
            else:
                table_father.display(self, "见证人（签名）：日期正确", "green")

        pattern = r".*[\S\s\n](.*年.*月.*日*)"
        timex = re.findall(pattern, self.contract_text)[-1]
        time = re.findall(pattern, self.contract_text)[-1]
        time = tyh.changeDate(time).replace(" ", "")
        if tyh.get_strtime_with_(time) == False:
            table_father.display(self, "落款时间：落款时间格式错误", 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, timex, '落款时间格式错误')
        else:
            table_father.display(self, "落款时间：落款时间格式正确", 'green')


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    if "抽样取证物品清单_.docx" in list:
        ioc = table9(my_prefix, my_prefix)
        contract_file_path = my_prefix + "抽样取证物品清单_.docx"
        ioc.check(contract_file_path, "抽样取证物品清单_.docx")
