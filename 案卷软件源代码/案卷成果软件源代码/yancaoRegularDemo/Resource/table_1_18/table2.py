import time

from yancaoRegularDemo.Resource.ReadFile import DocxData
from yancaoRegularDemo.Resource.tools.EntityRecognition import EntityRecognition
from yancaoRegularDemo.Resource.tools.tangyuhao_readfile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *

function_description_dict = {
    'head': '请核对文书编号（如川烟立XX号）',
    'reasonRight': '1.案由不为空，主观审查;2.案由的格式应为“涉嫌XXXX”;'
                   '3.违法行为性质表述应与法律条款保持一致;4.案由应与卷宗封面、行政处罚决定书、证据先行登记保存通知书、委托鉴定告知书、询问（调查）通知书等文书保持一致',
    'sourceRIGHT': '案件来源不应为空',
    'timeRight': '1.案发时间应具体到XX年XX月XX日XX时XX分;2.案发时间应与检查（勘验）笔录开始检查时间，以及案件调查终结报告等文书中所记录的日期保持一致',
    'placeRight': '1.案发地点应具体到XX市XX县（区）XX乡镇（街道）XX号;2.案发地点应与检查（勘验）笔录记载的的被检查（勘验）地点保持一致;'
                  '3.案发地点应与检查（勘验）笔录中检查（勘验）地点、询问（调查）通知书记载的地点保持一致',
    'basicInfOfPeople': '1.当事人姓名应与行政处罚决定书、检查（勘验）笔录、证据先行登记保存批准书等文书所记载的当事人名称（姓名）保持一致;'
                        '2.证件类型及号码应按照“身份证号：XXXX”格式填写，当事人为法人或其他组织的，应填写法定代表人或其他组织负责人的身份证号码。身份证号码应与其余文书所记载的当事人身份证号码保持一致;'
                        '3.当事人为自然人，地址一栏应填写身份证地址；当事人为法人或其他组织（如何判断法人与自然人），地址一栏应填写经营场所地址或注册地址。地址应与行政处罚决定书所记载的当事人住址、证据复制（提取）单当事人身份证记载地址、询问笔录当事人住址、案件处理审批表住址、结案报告表地址所记载的内容保持一致',
    'summaryOfTheCase_case1': '1.案情摘要应包含案发日期、执法主体、检查方式、检查地点、涉案物品等内容（是否会有标注）;'
                              '2.摘要中案发日期应与案发时间栏日期保持一致;'
                              '3.执法人员应写清参与执法的所有主体，并与勘验（检查）笔录记载的执法主体保持一致',
    'opinionsOfTheUndertaker': '1.应写明当事人涉嫌违反的条款、具体涉嫌的违法行为，主观审查;'
                               '2.应准确适用立案依据，主观审查;'
                               '3.承办人签名应有两名承办人员签名并注明日期，日期应在案发时间后7日内，并不得在案件调查终结报告处理意见栏所注明的日期之后',
    'opinionsOfTheDepartment': '1.承办部门负责人应作出是否同意立案或不予立案的意见;'
                               '2.承办部门负责人应签名并注明日期，日期应在案发时间后7日内，并不得在案件调查终结报告处理意见栏所注明的日期之后',
    'opinionsOfPersonInCharge': '1.承办案件的烟草专卖行政主管部门负责人签署审批意见栏不为空;'
                                '2. 承办案件的烟草专卖行政主管部门负责人应签名并注明日期，日期应在案发时间后7日内，并不得在案件调查终结报告处理意见栏所注明的日期之后',
}


class table2(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix
        self.entityrecognition = EntityRecognition()

        self.all_to_check = [
            "self.head()",
            "self.reasonRight()",
            "self.sourceRIGHT()",
            "self.timeRight()",
            "self.placeRight()",
            "self.basicInfOfPeople()",
            "self.summaryOfTheCase_case1()",
            "self.opinionsOfTheUndertaker()",
            "self.opinionsOfTheDepartment()",
            "self.opinionsOfPersonInCharge()",
        ]

    def get_strtime(self, text):
        text = text.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip()
        text = re.sub("\s+", " ", text)
        t = ""
        regex_list = [
            # # 2013年8月15日 22:46:21
            # "(\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}:\d{1,2})",
            # # "2013年8月15日 22:46"
            # "(\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2})",
            # "2014年5月11日"
            "(\d{4}-\d{1,2}-\d{1,2})",

        ]
        for regex in regex_list:
            t = re.search(regex, text)
            if t:
                t = t.group(1)
                return t
        else:
            return False

    def checkTime(self):
        # if os.path.exists(self.source_prifix + "检查（勘验）笔录_.docx") == 0:
        if tyh.file_exists(self.source_prifix, "检查（勘验）笔录") == False:
            table_father.display(self, "检查（勘验）时间：《勘验（检查）笔录》不存在", "red")
            return ""
        else:
            data1 = tyh.file_exists_open(self.source_prifix, "检查（勘验）笔录", file_1)
            text = data1.text
            time0 = re.findall("检查（勘验）时间：(.*?)\n", text)
            if time0 != [''] and time0 != []:
                time0 = time0[0].replace(" ", "")
                if "至" in time0:
                    time = re.findall("(.*)至.*?", time0)[0]
                    return time
                else:
                    return time0
            else:
                return ""

    def head(self):
        text = self.contract_text
        if "烟立" in text:
            pattern = r".烟立[[](.*?)[]].*"
        if "烟不立" in text:
            pattern = r".烟不立[[](.*?)[]].*"
        list = re.findall(pattern, text)
        pattern1 = r".*第(.*)号.*"
        list1 = re.findall(pattern1, text)
        if list == [] or list1 == [] or list[0].strip() == "" or list1[0].strip() == "":
            table_father.display(self, "表头：表头不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '新烟立', '表头不能为空')

    def checkPlace(self):
        # if os.path.exists(self.source_prifix + "检查（勘验）笔录_.docx") == 0:
        if tyh.file_exists(self.source_prifix, "检查（勘验）笔录") == False:
            table_father.display(self, "检查（勘验）地点：《勘验（检查）笔录》不存在", "red")
            return ""
        data1 = tyh.file_exists_open(self.source_prifix, "检查（勘验）笔录", file_1)
        text = data1.text
        time0 = re.findall("检查（勘验）地点：(.*?)\n", text)
        if time0 != [''] and time0 != []:
            time0 = time0[0].replace(" ", "")
            return time0
        else:
            return ""

    def reasonRight(self):
        text = self.contract_tables_content["案由"]
        text1, text2, text3, text4, text5 = "", "", "", "", ""
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "案由：案由为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案    由', '“案由”为空')
        else:
            table_father.display(self, "案由：案由不为空,主观审查", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, '案    由', '“案由”不为空,主观审查')
            pattern = re.compile(r'涉嫌.*')

            if re.match(pattern, text) != None:
                table_father.display(self, "案由：案由格式正确", "green")
            else:
                table_father.display(self, "案由：案由未以”涉嫌“开头", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案    由', '“案由”未以”涉嫌“开头')
            flag = 0
            # if os.path.exists(self.source_prifix + "卷宗封面_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "卷宗封面"):
                data1 = tyh.file_exists_open(self.source_prifix, "卷宗封面", file_1)
                text1 = data1.tabels_content["案由"].replace(" ", "")

            # if os.path.exists(self.source_prifix + "行政处罚决定书_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "行政处罚决定书"):
                data1 = tyh.file_exists_open(self.source_prifix, "行政处罚决定书", file_1)
                text0 = data1.text
                texta = re.findall(".*案由：(.*?)\n", text0)
                if texta != [''] and texta != []:
                    text2 = texta[0].replace(" ", "")

            # if os.path.exists(self.source_prifix + "证据先行登记保存批准书_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "证据先行登记保存批准书"):
                data1 = tyh.file_exists_open(self.source_prifix, "证据先行登记保存批准书", file_1)
                text0 = data1.text
                texta = re.findall(".*涉嫌(.*?)行为", text0)
                if texta != [''] and texta != []:
                    text3 = texta[0].replace(" ", "")

            # if os.path.exists(self.source_prifix + "委托鉴定告知书_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "委托鉴定告知书"):
                data1 = tyh.file_exists_open(self.source_prifix, "委托鉴定告知书", file_1)
                text0 = data1.text
                texta = re.findall(".*我局在调查你涉嫌(.*?)一案中", text0)
                if texta != [''] and texta != []:
                    text4 = texta[0].replace(" ", "")

            # if os.path.exists(self.source_prifix + "询问（调查）通知书_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "询问（调查）通知书"):
                data1 = tyh.file_exists_open(self.source_prifix, "询问（调查）通知书", file_1)
                text0 = data1.text
                texta = re.findall(".*涉嫌(.*?)一案", text0)
                if texta != [''] and texta != []:
                    text5 = texta[0].replace(" ", "")

            text_all = [text1, text2, text3, text4, text5]

            for i in range(1, 5):
                if text_all[i] != "" and text_all[i] not in text:
                    flag = 1
            if flag:
                table_father.display(self, "案由：“案由”没有与卷宗封面（" + str(text1) + "）、行政处罚决定书（" + str(
                    text2) + "）、证据先行登记保存通知书（" + str(text3) + "）、委托鉴定告知书（" + str(
                    text4) + "）、询问（调查）通知书（" + str(text5) + "）等文书保持一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案    由',
                                   "“案由”没有与卷宗封面（" + str(text1) + "）、行政处罚决定书（" + str(
                                       text2) + "）、证据先行登记保存通知书（" + str(text3) + "）、委托鉴定告知书（" + str(
                                       text4) + "）、询问（调查）通知书（" + str(text5) + "）等文书保持一致")
            else:
                table_father.display(self,
                                     "案由：案由与卷宗封面、行政处罚决定书、证据先行登记保存通知书、委托鉴定告知书、询问（调查）通知书等文书保持一致",
                                     "green")

    def sourceRIGHT(self):
        text = self.contract_tables_content["案件来源"]
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "案件来源：案件来源为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案件来源', '案件来源为空')
        else:
            table_father.display(self, "案件来源：案件来源不为空,主观审查", "green")

    def timeRight(self):
        text = self.contract_tables_content["案发时间"]
        str1 = self.checkTime()
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "案发时间：缺少案发时间", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案发时间', '缺少案发时间')
            return
        if str1.strip() == "" or str1.strip() == "/":
            table_father.display(self, "案发时间：缺少检查（勘验）笔录开始检查时间", "red")
            return
        if text == str1:
            table_father.display(self,
                                 "案发时间：案发时间与检查（勘验）笔录开始检查时间，以及案件调查终结报告等文书中所记录的日期一致",
                                 "green")
        else:
            table_father.display(self, "案发时间：案发时间（" + str(
                text) + "）与检查（勘验）笔录开始检查时间，以及案件调查终结报告等文书中所记录的日期（" + str(
                str1) + "）不一致", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案发时间', "”案发时间“（" + str(
                text) + "）与检查（勘验）笔录开始检查时间，以及案件调查终结报告等文书中所记录的日期（" + str(
                str1) + "）不一致")

    def placeRight(self):
        text = self.contract_tables_content["案发地点"]
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "案发地点：“立案（不予立案）报告表_案发地点”不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案发地点', '“立案（不予立案）报告表_案发地点”不能为空')
            return
        pattern = re.compile(r'.*市.*[县|区].*[乡镇|街道|街|镇|乡|道].*号.*')
        if re.search(pattern, text) == None:
            table_father.display(self, "案发地点：“案发地点”格式不正确，应具体到XX市XX县（区）XX乡镇（街道）XX号", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案发地点',
                               '“案发地点”格式不正确，应具体到XX市XX县（区）XX乡镇（街道）XX号')
        place = self.checkPlace()
        if place.strip() == "" or place.strip() == "/":
            table_father.display(self, "案发地点：缺少检查（勘验）笔录地点", "red")
        else:
            tmp_str = place.split()  # tmp_str = ['a' ,'b' ,'c']
            place0 = ''.join(tmp_str)  # 用一个空字符串join列表
            if place0.replace(" ", "") == text.replace(" ", ""):
                table_father.display(self, "案发地点：案发地点应与检查（勘验）笔录记载的的被检查（勘验）地点保持一致",
                                     "green")
            else:
                table_father.display(self, "案发地点：案发地点（" + str(
                    text) + "）应与检查（勘验）笔录记载的的被检查（勘验）地点（" + str(place0) + "）不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案发地点',
                                   "案发地点（" + str(text) + "）应与检查（勘验）笔录记载地点（" + str(place0) + "）不一致")

        if tyh.file_exists(self.source_prifix, "询问（调查）通知书") == 0:
            table_father.display(self, "案发地点：《勘验（检查）笔录》不存在", "red")
        else:
            data1 = tyh.file_exists_open(self.source_prifix, "询问（调查）通知书", file_1)
            text0 = data1.text
            pattern = r'.*在(.*)查获你.*'
            place = re.findall(pattern, text0)
            if place == [] or place[0].replace(" ", "") == "":
                table_father.display(self, "案发地点：询问（调查）通知书_ 地点获取失败", "red")
            else:
                place = place[0].replace(" ", "")
                if place == text:
                    table_father.display(self, "案发地点：案发地点与询问（调查）通知书记载的的被检查（勘验）地点一致",
                                         "green")
                else:
                    table_father.display(self, "案发地点：案发地点（" + str(
                        text) + "）应与询问（调查）通知书记载的的被检查（勘验）地点（" + str(place) + "）不一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '案发地点',
                                       "案发地点（" + str(text) + "）应与询问（调查）通知书记载的的被检查（勘验）地点（" + str(
                                           place) + "）不一致")

    def basicInfOfPeople(self):
        p1, p2, p3, p4, p5 = "", "", "", "", ""
        text = self.contract_tables_content["当事人"]
        text1, text2, text3 = "", "", ""
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "当事人：立案（不予立案）报告表_当事人不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '当 事 人', '“当事人”不能为空')
        else:
            if tyh.file_exists(self.source_prifix, "行政处罚决定书"):
                data1 = tyh.file_exists_open(self.source_prifix, "行政处罚决定书", file_1)
                text0 = data1.text
                texta = re.findall(".*当事人：(.*?)\n", text0)
                p0 = re.findall(".*住址：(.*?)\n", text0)
                if texta != [''] and texta != []:
                    text1 = texta[0].replace(" ", "")
                if p0 != [''] and p0 != []:
                    p1 = p0[0].replace(" ", "")

            # if os.path.exists(self.source_prifix + "检查（勘验）笔录_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "检查（勘验）笔录"):
                data1 = tyh.file_exists_open(self.source_prifix, "检查（勘验）笔录", file_1)
                text0 = data1.text
                texta = re.findall(".*被检查（勘验）人姓名：(.*?)[\n|性别：]", text0)
                if texta != [''] and texta != []:
                    text2 = texta[0].replace(" ", "")

            # if os.path.exists(self.source_prifix + "证据先行登记保存批准书_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "证据先行登记保存批准书"):
                data1 = tyh.file_exists_open(self.source_prifix, "证据先行登记保存批准书", file_1)
                text0 = data1.text
                texta = re.findall(".*因(.*?)涉嫌", text0)
                if texta != [''] and texta != []:
                    text3 = texta[0].replace(" ", "")

            list1 = [text1, text2, text3]
            flag = 0
            for i, value in enumerate(list1):
                if value != "" and value != text:
                    flag = 1
            if flag:
                table_father.display(self, "当事人：当事人姓名（" + str(
                    text) + "）与行政处罚决定书、检查（勘验）笔录、证据先行登记保存批准书等文书所记载的当事人名称（姓名）（" + str(
                    list1[0]) + str(list1[1]) + str(list1[2]) + "）不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '当 事 人', "当事人姓名（" + str(
                    text) + "）与行政处罚决定书、检查（勘验）笔录、证据先行登记保存批准书等文书所记载的当事人名称（姓名）（" + str(
                    list1[0]) + str(list1[1]) + str(list1[2]) + "）不一致")
            else:
                table_father.display(self,
                                     "当事人：当事人姓名与行政处罚决定书、检查（勘验）笔录、证据先行登记保存批准书等文书所记载的当事人名称（姓名）一致",
                                     "green")

        text1, text2 = "", ""
        text = self.contract_tables_content["证件类型及号码"]
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "证件类型及号码：立案（不予立案）报告表_证件类型及号码 不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '证件类型及号码', '证件类型及号码 不能为空')
        else:
            # if os.path.exists(self.source_prifix + "调查终结报告_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "调查终结报告"):
                data1 = tyh.file_exists_open(self.source_prifix, "调查终结报告", file_1)
                text0 = data1.tabels_content["证件类型及号码"]
                if text0 != "":
                    text1 = text0.replace(" ", "")

            # if os.path.exists(self.source_prifix + "案件处理审批表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "案件处理审批表"):
                data1 = tyh.file_exists_open(self.source_prifix, "案件处理审批表", DocxData)
                # text0 = data1.tabels_content["证件类型号码"]
                # if text0 != "":
                #     text2 = text0.replace(" ","")
                try:
                    p0 = data1.tabels_content["住址"]
                    if p0 != "":
                        p3 = p0
                except Exception as e:
                    table_father.display(self, "住址 has occurred an error:" + str(e.args))

            list1 = [text1, text2]
            if "身份证号：" not in text:
                table_father.display(self, "证件类型及号码：证件类型及号码格式不正确，应为 ”身份证号：XXXX“", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '证件类型及号码', '证件类型及号码格式不正确，应为 ”身份证号：XXXX“')
            else:
                flag = 0
                for i, val in enumerate(list1):
                    pattern = re.compile(r'.*：(.*)')
                    if re.findall(pattern, val) != re.findall(pattern, text):
                        flag = 1
                if flag:
                    table_father.display(self, "证件类型及号码：身份证号码（" + str(
                        text) + "）与其余文书所记载的当事人身份证号码（" + str(list1[0]) + str(list1[1]) + "）不一致",
                                         "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '证件类型及号码',
                                       "身份证号码（" + str(text) + "）与其余文书所记载的当事人身份证号码（" + str(
                                           list1[0]) + str(list1[1]) + "）不一致")
                else:
                    table_father.display(self, "证件类型及号码：身份证号码与其余文书所记载的当事人身份证号码一致",
                                         "green")

                if tyh.is_id_number((text[5:])) == False:
                    table_father.display(self, "证件类型及号码：身份证号码不正确", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '证件类型及号码', '身份证号码不正确')

        flag = 0
        text = self.contract_tables_content["地址"]
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "地址：立案（不予立案）报告表_地址 不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '地    址', '地址 不能为空')
        else:
            # if os.path.exists(self.source_prifix + "结案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "结案报告表"):
                data1 = tyh.file_exists_open(self.source_prifix, "结案报告表", file_1)
                text0 = data1.tabels_content["地址"]
                if text0 != "":
                    p4 = text0.replace(" ", "")

            list1 = [p1, p3, p4]
            for i, val in enumerate(list1):
                if val != "" and text != val:
                    flag = 1
            if flag:
                table_father.display(self,
                                     "地址：地址（" + str(
                                         text) + "）与行政处罚决定书所记载的当事人住址、证据复制（提取）单当事人身份证记载地址、询问笔录当事人住址、案件处理审批表住址、结案报告表地址所记载的内容（" + str(
                                         list1[0]) + str(list1[1]) + str(list1[2]) + "）没有保持一致",
                                     "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '地    址',
                                   "地址（" + str(
                                       text) + "）与行政处罚决定书所记载的当事人住址、证据复制（提取）单当事人身份证记载地址、询问笔录当事人住址、案件处理审批表住址、结案报告表地址所记载的内容（" + str(
                                       list1[0]) + str(list1[1]) + str(list1[2]) + "）没有保持一致")

    # def summaryOfTheCase(self):
    #     text = self.contract_tables_content["案情摘要"]
    #     if text.strip() == "" or text.strip() == "/":
    #         table_father.display(self, "案情摘要：立案（不予立案）报告表_案情摘要 不能为空", "red")
    #         tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要', '“案情摘要”不能为空')
    #         return
    #     time = self.get_strtime(text).split("-")
    #     for i in range(0, len(time)):
    #         time[i] = int(time[i])
    #
    #     time1 = self.get_strtime(self.contract_tables_content["案发时间"]).split("-")
    #     for i in range(0, len(time1)):
    #         time1[i] = int(time1[i])
    #
    #     if time != time1:
    #         table_father.display(self, "案情摘要：摘要中案发日期（"+str(time)+"）应与案发时间栏日期（"+str(time1)+"）没有保持一致", "red")
    #         tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要', "摘要中案发日期（"+str(time)+"）应与案发时间栏日期（"+str(time1)+"）没有保持一致")
    #     else:
    #         table_father.display(self, "案情摘要：摘要中案发日期应与案发时间栏日期一致", "green")

    def summaryOfTheCase_case1(self):
        text = self.contract_tables_content["案情摘要"]
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "案情摘要：立案（不予立案）报告表_案情摘要 不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要', '案情摘要 不能为空')
            return
        time = self.get_strtime(text)

        # 日期与本表头中检查（勘验）时间的开始时间一致
        if time != False:
            time = time.split("-")
            for i in range(0, len(time)):
                time[i] = int(time[i])

            time1 = self.get_strtime(self.contract_tables_content["案发时间"]).split("-")
            for i in range(0, len(time1)):
                time1[i] = int(time1[i])

            if time != time1:
                table_father.display(self, "案情摘要：摘要中案发日期（" + str(time) + "）应与案发时间栏日期（" + str(
                    time1) + "）没有保持一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要',
                                   "摘要中案发日期（" + str(time) + "）应与案发时间栏日期（" + str(
                                       time1) + "）没有保持一致")
            else:
                table_father.display(self, "案情摘要：摘要中案发日期应与案发时间栏日期一致", "green")

            jiancha_time = tyh.jiancha_time(self.source_prifix)
            if jiancha_time != False:
                s_time = jiancha_time[0][0]
                s_time = tyh.get_strtime_5(s_time)
                # print(s_time)
                if s_time != False:
                    time = self.get_strtime(text)
                    # print(time)
                    if time in s_time:
                        table_father.display(self, "案情摘要：摘要中案发日期应与 检查勘验日期 保持一致", "green")
                    else:
                        table_father.display(self, "案情摘要：摘要中案发日期（" + str(time) + "）应与检查勘验日期（" + str(
                            s_time) + "）没有保持一致", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要',
                                           "摘要中案发日期（" + str(time) + "）应与检查勘验日期（" + str(
                                               s_time) + "）没有保持一致")

                else:
                    table_father.display(self, "案情摘要：检查勘验 日期格式错误", "red")
            else:
                table_father.display(self, "案情摘要：检查勘验 日期无法找到", "red")
        else:
            table_father.display(self, "案情摘要：摘要中无案发日期或日期格式错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要', '摘要中无案发日期或日期格式错误')

        # 案情摘要文本中应包含执法人员
        text = self.contract_tables_content["案情摘要"]
        if "执法人员" not in text:
            table_father.display(self, "案情摘要：“案情摘要”未包含执法人员", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要', '“案情摘要”未包含执法人员')
        else:
            table_father.display(self, "案情摘要：案情摘要 包含执法人员", "green")

        # 案情摘要文本中地址信息应与检查（勘察）中地点一致
        text = self.contract_tables_content["案情摘要"]
        place = self.entityrecognition.get_identity_with_tag(text, 'LOC')
        # print(place)
        jiancha_place = tyh.jiancha_place(self.source_prifix)
        if jiancha_place != False:
            jiancha_place = jiancha_place[0]
            if tyh.list_str_match(place, jiancha_place) == False:
                table_father.display(self, "案情摘要：案情摘要未包含地点 或者 与检查勘验地点（" + str(
                    jiancha_place) + "）不匹配", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要',
                                   "案情摘要未包含地点 或者 与检查勘验地点（" + str(jiancha_place) + "）不匹配")
            else:
                table_father.display(self, "案情摘要：案情摘要地点与检查勘验地点匹配", "green")

        else:
            table_father.display(self, "案情摘要：无法得到 检查勘验地点", "red")

        # 文本中是否包含烟草专卖零售许可证、卷烟数目、法律条款等要素，若无则提示
        text = self.contract_tables_content["案情摘要"]
        if "许可证" not in text or "持证" not in text or '烟草专卖许可证' not in text:
            table_father.display(self, "案情摘要：案情摘要未包含 烟草专卖零售许可证要素", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要', '案情摘要未包含 烟草专卖零售许可证要素')
        else:
            table_father.display(self, "案情摘要：案情摘要包含 烟草专卖零售许可证要素", "green")

        pattern = ".*共计(.*)个品种(.*)条.*"
        if re.match(pattern, text) == None:
            table_father.display(self, "案情摘要：案情摘要未包含 卷烟数目 要素", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要', '案情摘要未包含 卷烟数目 要素')
        else:
            table_father.display(self, "案情摘要：案情摘要包含 卷烟数目要素", "green")

        # pattern = ".*(《.*》).*"
        # if re.match(pattern, text) == None:
        #     table_father.display(self, "案情摘要未包含 法律条款 要素", "red")
        #     tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要', '案情摘要未包含 法律条款 要素')
        # else:
        #     table_father.display(self, "案情摘要包含 法律条款要素", "green")

        if "凭证" not in text and "烟草专卖品准运证" not in text and "有效证明" not in text:
            table_father.display(self, "案情摘要：案情摘要未包含 法律条款 要素", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案情摘要', '案情摘要未包含 法律条款 要素')
        else:
            table_father.display(self, "案情摘要：案情摘要包含 法律条款要素", "green")

    def opinionsOfTheUndertaker(self):
        time0 = False
        text = self.contract_tables_content["承办人意见"]
        pattern = re.compile(r'(.*)签名.*')
        yijian = re.findall(pattern, text)
        if yijian == [''] or yijian == [] or yijian[0].strip() == "":
            table_father.display(self, "承办人意见：承办人意见不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办人\r意见', '承办人意见 不能为空')
        sign, date = tyh.sign_date(text)
        if sign == [''] or sign == [] or sign[0].strip() == "":
            table_father.display(self, "承办人意见：承办人意见未签名", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办人\r意见', '承办人意见未签名')
        if date == [''] or date == [] or date[0].strip() == "":
            table_father.display(self, "承办人意见：承办人意见未注明日期", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办人\r意见', '承办人意见未注明日期')
        else:
            date = date[0].replace(" ", "")
            time1 = self.get_strtime(self.contract_tables_content["案发时间"]).replace(" ", "")
            time0 = self.get_strtime(date).replace(" ", "")
            if time1 and time0:
                if tyh.time_differ(time0, time1) >= 7:
                    table_father.display(self, "承办人意见：承办部门意见时间（" + str(time0) + "）应在案发时间（" + str(
                        time1) + "）后7日内", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '承办人\r意见',
                                       "承办部门意见时间（" + str(time0) + "）应在案发时间（" + str(time1) + "）后7日内")
            else:
                table_father.display(self, "承办人意见：承办部门意见时间或者案发时间未能提取", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办人\r意见', '承办部门意见时间或者案发时间未能提取')
        d_text = ""

        # if os.path.exists(self.source_prifix + "调查终结报告_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "调查终结报告"):
            data1 = tyh.file_exists_open(self.source_prifix, "调查终结报告", file_1)
            text0 = data1.tabels_content["处理意见"]
            if text0 != "":
                d_text = text0.replace(" ", "")

        _, d_date = tyh.sign_date(d_text)
        if d_date == [""] or d_date == [] or d_date[0].strip() == "":
            table_father.display(self, "承办人意见：案件调查终结报告处理意见栏未注明日期", "red")
        else:
            date = d_date[0]
            d_time = self.get_strtime(date)
            if time0 and d_time:
                if tyh.time_differ(time0, d_time) > 0:
                    table_father.display(self, "承办人意见：承办人时间（" + str(
                        time0) + "）不得在案件调查终结报告处理意见栏所注明的日期（" + str(d_time) + "）之后", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '承办人\r意见', '承办人时间（' + str(
                        time0) + "）不得在案件调查终结报告处理意见栏所注明的日期（" + str(d_time) + '）之后')
                else:
                    table_father.display(self, "承办人意见：承办人时间在案件调查终结报告处理意见栏所注明的日期之前",
                                         "green")
            else:
                table_father.display(self,
                                     "承办人意见：承办人意见时间或者案件调查终结报告处理意见栏所注明的日期未成功提取",
                                     "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办人\r意见',
                                   '承办人意见时间或者案件调查终结报告处理意见栏所注明的日期未成功提取')

    def opinionsOfTheDepartment(self):
        time0 = False
        text = self.contract_tables_content["承办部门意见"]
        pattern = re.compile(r'(.*)签名.*')
        yijian = re.findall(pattern, text)
        if yijian == [''] or yijian == [] or yijian[0].strip() == "":
            table_father.display(self, "承办部门意见：立案（不予立案）报告表_承办部门意见 不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '立案（不予立案）报告表_承办部门意见 不能为空')
            return
        if "同意立案" in text:
            table_father.display(self, "承办部门意见：同意立案", "green")
        elif "不" in text and "立案" in text:
            table_father.display(self, "承办部门意见：不予立案", "green")
        else:
            table_father.display(self, "承办部门意见：未清楚表明是否立案", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '未清楚表明是否立案')
        sign, date = tyh.sign_date(text)
        if sign == [''] or sign == [] or sign[0].strip() == "":
            table_father.display(self, "承办部门意见：承办部门意见未签名", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '承办部门意见未签名')
        if date == [""] or date == [] or date[0].strip() == "":
            table_father.display(self, "承办部门意见：承办人意见未注明日期", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '承办人意见未注明日期')
        else:
            date = date[0]
            time1 = self.get_strtime(self.contract_tables_content["案发时间"])
            time0 = self.get_strtime(date)
            if time1 and time0:
                if tyh.time_differ(time0, time1) >= 7:
                    table_father.display(self, "承办部门意见：承办部门意见时间（" + str(time0) + "）应在案发时间（" + str(
                        time1) + "）后7日内", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见',
                                       "承办部门意见时间（" + str(time0) + "）应在案发时间（" + str(time1) + "）后7日内")
            else:
                table_father.display(self, "承办部门意见：承办部门意见时间或者案发时间未能提取", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '承办部门意见时间或者案发时间未能提取')

        d_text = ""
        # if os.path.exists(self.source_prifix + "调查终结报告_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "调查终结报告"):
            data1 = tyh.file_exists_open(self.source_prifix, "调查终结报告", file_1)
            text0 = data1.tabels_content["处理意见"]
            if text0 != "":
                d_text = text0
        _, d_date = tyh.sign_date(d_text)
        if d_date == [""] or d_date == [] or d_date[0].strip() == "":
            table_father.display(self, "承办部门意见：案件调查终结报告处理意见栏未注明日期", "red")
        else:
            date = d_date[0]
            d_time = self.get_strtime(date)
            if time0 and d_time:
                if tyh.time_differ(time0, d_time) > 0:
                    table_father.display(self, "承办部门意见：承办部门时间（" + str(
                        time0) + "）不得在案件调查终结报告处理意见栏所注明的日期（" + str(d_time) + "）之后", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', "承办部门时间（" + str(
                        time0) + "）不得在案件调查终结报告处理意见栏所注明的日期（" + str(d_time) + "）之后")
                else:
                    table_father.display(self, "承办部门意见：承办部门时间在案件调查终结报告处理意见栏所注明的日期之前",
                                         "green")
            else:
                table_father.display(self,
                                     "承办部门意见：承办部门意见时间或者案件调查终结报告处理意见栏所注明的日期未成功提取",
                                     "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见',
                                   '承办部门意见时间或者案件调查终结报告处理意见栏所注明的日期未成功提取')

    def opinionsOfPersonInCharge(self):
        text = self.contract_tables_content["负责人意见"]
        time0 = ""
        pattern = re.compile(r'(.*)签名.*')
        yijian = re.findall(pattern, text)
        if yijian == [''] or yijian == [] or yijian[0].strip() == "":
            table_father.display(self, "负责人意见：负责人意见不为空 不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人\r意见', '负责人意见不为空 不能为空')
            return
        sign, date = tyh.sign_date(text)
        if sign == [''] or sign == [] or sign[0].strip() == "":
            table_father.display(self, "负责人意见：负责人意见未签名", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人\r意见', '负责人意见未签名')
        if date == [""] or date == [] or date[0].strip() == "":
            table_father.display(self, "负责人意见：负责人意见未注明日期", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人\r意见', '负责人意见未注明日期')
        else:
            date = date[0]
            time1 = self.get_strtime(self.contract_tables_content["案发时间"])
            time0 = self.get_strtime(date)
            if tyh.time_differ(time0, time1) >= 7:
                table_father.display(self, "负责人意见：负责人意见时间（" + str(time0) + "）应在案发时间（" + str(
                    time1) + "）后7日内", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '负责人\r意见',
                                   "负责人意见时间（" + str(time0) + "）应在案发时间（" + str(time1) + "）后7日内")

        d_text = ""
        if tyh.file_exists(self.source_prifix, "调查终结报告"):
            data1 = tyh.file_exists_open(self.source_prifix, "调查终结报告", file_1)
            text0 = data1.tabels_content["处理意见"]
            if text0 != "":
                d_text = text0
        _, d_date = tyh.sign_date(d_text)
        if d_date == [""] or d_date == [] or d_date[0].strip() == "":
            table_father.display(self, "负责人意见：案件调查终结报告处理意见栏未注明日期", "red")
        else:
            date = d_date[0]
            d_time = self.get_strtime(date)
            if date and d_time:
                if tyh.time_differ(time0, d_time) > 0:
                    table_father.display(self, "负责人意见：负责人意见时间（" + str(
                        time0) + "）不得在案件调查终结报告处理意见栏所注明的日期（" + str(d_time) + "）之后", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "负责人\r意见', '负责人意见时间（" + str(
                        time0) + "）不得在案件调查终结报告处理意见栏所注明的日期（" + str(d_time) + "）之后")
                else:
                    table_father.display(self, "负责人意见：负责人意见时间在案件调查终结报告处理意见栏所注明的日期之前",
                                         "green")
            else:
                table_father.display(self,
                                     "负责人意见：负责人意见时间或者案件调查终结报告处理意见栏所注明的日期未成功提取",
                                     "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '负责人\r意见',
                                   "负责人意见时间或者案件调查终结报告处理意见栏所注明的日期未成功提取")

    # 用于MultiTableProcessor，获取案由
    def get_cause_of_action(self, contract_file_path):
        # if os.path.exists(contract_file_path) == True:
        if tyh.file_exists(self.source_prifix, contract_file_path):
            data = tyh.file_exists_open(self.source_prifix, contract_file_path, file_1)
            return data.tabels_content["案由"]
        else:
            return False

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
        self.mw = win32com.client.Dispatch("Word.Application")
        self.doc = self.mw.Documents.Open(self.my_prefix + file_name_real)
        print(self.my_prefix + file_name_real)
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


if __name__ == '__main__':
    # ioc = table3()
    # contract_file_path = my_prefix + "延长立案期限审批表_.docx"
    # ioc.check(contract_file_path)
    my_prefix = "C:\\Users\\Zero\\OneDrive\\案卷\\tyh\\"
    list = os.listdir(my_prefix)
    if "立案报告表_.docx" in list:
        ioc = table2(my_prefix, my_prefix)
        contract_file_path = my_prefix + "立案报告表_.docx"
        ioc.check(contract_file_path, "立案报告表_.docx")
