from yancaoRegularDemo.Resource.tools.tangyuhao_readfile import *

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *
from yancaoRegularDemo.Resource.tools.simple_content import Simple_Content

from yancaoRegularDemo.Resource.tools.utils import is_valid_date

import win32com.client

from yancaoRegularDemo.Resource.ReadFile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

from warnings import simplefilter

function_description_dict = {
    'timeRight': '1.举报时间应具体到XX年XX月XX日XX时XX分；2.举报时间应在调查取证阶段中《勘验（检查）笔录》文书中涉及的检查开始时间30分钟之前。',
    'formRIGHT': '填写内容是否包含“举报”二字。',
    'basicInfOfPeople': '1.是否填写举报人的姓名、性别、住址等基本情况；2.如举报人不愿留下姓名或要求保密，应在“举报人有关情况”栏中填写“举报人要求保密”。',
    'reportContent': '1.接收举报的日期是否具体到XX年XX月XX日，应与举报时间栏填写日期相一致；2.被举报人是否有具体的姓名，如无具体姓名，应填写相关的手机号码、车牌号码、违法地点等。',
    'ReceptionPersonOpinions': '1.意见栏不为空，提出意见人应当签名，并注明日期。2.举报人递交书面举报材料的，举报材料应附在举报记录表后面。'
                               '3.接待人（承办部门负责人）是否填写意见并签名。4.接待人意见（承办部门负责人意见）栏应注明日期，并应与举报时间栏填写日期保持一致',
}


class table1(table_father):
    def __init__(self, my_prefix, source_prifix):
        # super(table1, self).__init__()
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix

        self.all_to_check = [
            "self.timeRight()",
            "self.formRIGHT()",
            "self.basicInfOfPeople()",
            "self.reportContent()",
            "self.ReceptionPersonOpinions()"
        ]

    def timeRight(self):
        ttime = self.contract_tables_content["举报时间"]
        if ttime.strip() == "" or ttime.strip() == "/":
            table_father.display(self, "举报时间：举报记录表_举报时间 不为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '举报时间', '举报时间不为空')
        else:
            if is_valid_date(ttime) == True:
                table_father.display(self, "举报时间：举报时间格式正确", "green")
            else:
                table_father.display(self, "举报时间：举报时间格式错误，格式应为XX年XX月XX日XX时XX分,文中为: " + ttime,
                                     "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '举报时间',
                                   "举报时间格式错误，格式应为XX年XX月XX日XX时XX分,文中为: " + ttime)
                return
            # time = re.findall("检查（勘验）时间：(.*?)\n", self.contract_text)[0]
            # if os.path.exists(self.source_prifix + "检查（勘验）笔录_.docx") == 0:
            if tyh.file_exists(self.source_prifix, "检查（勘验）笔录") == False:
                table_father.display(self, "举报时间：《勘验（检查）笔录》不存在", "red")

            else:
                data = tyh.file_exists_open(self.source_prifix, "检查（勘验）笔录", file_1)
                text = data.text
                time = re.findall("检查（勘验）时间：(.*?)\n", text)[0]
                if "至" in time:
                    time = re.findall("(.*)至.*?", time)[0]

                if time == None:
                    table_father.display(self, "举报时间：《勘验（检查）笔录》时间不存在", "red")
                else:
                    tmp_str = time.split()  # tmp_str = ['a' ,'b' ,'c']
                    str = ''.join(tmp_str)  # 用一个空字符串join列表
                    # print(str)
                    # print(time)
                    time0 = re.findall(r"\d+\.?\d*", time)  # 检查（勘验）时间
                    # print(time0)

                    time1 = re.findall(r"\d+\.?\d*", ttime)  # 举报时间

                    if int(time1[0]) > int(time0[0]) or int(time1[1]) > int(time0[1]) or int(time1[2]) > int(
                            time0[2]) or int(time1[3]) > int(time0[3]) or \
                            int(time1[3]) * 60 + \
                            int(time1[4]) - int(time0[3]) * 60 - int(time0[4]) >= -30:
                        table_father.display(self,
                                             "举报时间：举报时间应在调查取证阶段中《勘验（检查）笔录》文书中涉及的检查开始时间（" + time + "）30分钟之前",
                                             "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, '举报时间',
                                           "举报时间应在调查取证阶段中《勘验（检查）笔录》文书中涉及的检查开始时间（" + time + "）30分钟之前")
                        return
                    else:
                        table_father.display(self,
                                             "举报时间：举报时间在调查取证阶段中《勘验（检查）笔录》文书中涉及的检查开始时间（" + time + "）30分钟之前",
                                             "green")

    def formRIGHT(self):
        form = self.contract_tables_content["举报形式"]
        if form.strip() == "" or form.strip() == "/":
            table_father.display(self, "举报形式：举报记录表_举报形式 不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '举报形式', '举报形式 不能为空')
        else:
            pattern = re.compile(r'.*举报.*')
            if re.match(pattern, form) != None:
                table_father.display(self, "举报形式：举报形式正确", "green")
            else:
                table_father.display(self, "举报形式：举报形式未有“举报”字样", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '举报形式', '举报形式未有“举报”字样')

    def basicInfOfPeople(self):
        s = Simple_Content()
        text = self.contract_tables_content["举报人有关情况"]
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "举报人有关情况：举报记录表_举报人有关情况 不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '举报人\r有关情况', '举报人有关情况 不能为空')
        else:
            if "保密" in text:
                table_father.display(self, "举报人有关情况：举报人要求保密有关情况", "green")
            else:
                if "电话" in text or "电话：" in text:
                    flag = 1
                    pattern = r".*[电话：|电话](.*)[。,]"
                    list = re.findall(pattern, text)
                    if len(list[0]) != 11:
                        flag = 0
                    list1 = s.match_re(s.pattern_strings["phone_number"], text)
                    if flag == 0 or list1 == [''] or list1 == [] or list1[0][0].strip() == '':
                        table_father.display(self, "举报人有关情况：举报人有关情况 电话格式错误", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, list[0], '举报人有关情况 电话格式错误')

                if "姓名" not in text or "性别" not in text or "住址" not in text:
                    table_father.display(self, "举报人有关情况：未同时填写举报人的姓名、性别、住址等基本情况", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '举报人\r有关情况',
                                       '未同时填写举报人的姓名、性别、住址等基本情况')
                else:
                    table_father.display(self, "举报人有关情况：已填写举报人基本情况", "green")

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

    def reportContent(self):

        text = self.contract_tables_content["举报内容"]

        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "举报内容：举报记录表_举报内容 不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '举报内容', '举报内容 不能为空')
        else:
            pattern = re.compile('.*?\d{4}年\d{1,2}月\d{1,2}日.*')
            time = self.get_strtime(text)
            time0 = self.get_strtime(self.contract_tables_content["举报时间"])
            if time == False:
                table_father.display(self, "举报内容：举报内容时间格式错误或无时间", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '举报内容', '举报内容时间格式错误或无时间')
            elif time != time0:
                table_father.display(self, "举报内容：举报内容时间与举报时间栏填写日期不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '举报内容',
                                   "举报内容时间（" + time + "）与举报时间栏填写日期（" + time0 + "）不一致")
            else:
                table_father.display(self, "举报内容：举报内容时间正确", "green")
            s = Simple_Content()
            if "电话" in text or "电话：" in text:
                flag = 1
                pattern = r".*[电话：|电话](.*)[。,]"
                list = re.findall(pattern, text)
                print('list', list)
                print('list[0]', list[0])
                if len(list[0]) != 11:
                    print('***', list[0])
                    flag = 0
                list1 = s.match_re(s.pattern_strings["phone_number"], text)
                if flag == 0 or list1 == [''] or list1 == [] or list1[0][0].strip() == '':
                    table_father.display(self, "举报内容：”举报内容信息“电话格式错误", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, list[0], '”举报内容“电话格式错误')
            if "车牌号" in text or "车牌号：" in text or "车牌" in text or "车牌：" in text:
                flag = 1
                pattern = r".*[车牌号：|车牌号|车牌|车牌号：](.*)[。,]"
                list = re.findall(pattern, text)
                if len(list[0]) != 7:
                    flag = 0
                list2 = s.match_re(s.pattern_strings["license_plate_number"], text)
                if flag == 0 or list2 == [''] or list2 == [] or list2[0][0].strip() == '':
                    table_father.display(self, "举报内容：举报内容信息车牌号格式错误", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, list[0], '”举报内容“"车牌号格式错误')

    def ReceptionPersonOpinions(self):
        text = self.contract_tables_content["接待人意见"]
        time0 = self.contract_tables_content["举报时间"]
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "接待人意见：”举报记录表_接待人意见“不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '接待人意见', '”接待人意见“不能为空')
        else:
            pattern = re.compile(r'(.*)签名')
            yijian = re.findall(pattern, text)
            if yijian == [''] or yijian == [] or yijian[0].strip() == "":
                table_father.display(self, "接待人意见：接待人意见不能为空", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '接待人意见', '“接待人意见”不能为空')
            sign, date = tyh.sign_date(text)
            if sign == [''] or sign == [] or sign[0].strip() == "":
                table_father.display(self, "接待人意见：接待人未签名", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '接待人意见', '接待人未签名')

            if date == [""] or date == [] or date[0].strip() == "":
                table_father.display(self, "接待人意见：接待人未填写日期", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '接待人意见', '接待人未填写日期')
            else:
                date = date[0]
                if self.get_strtime(date) != self.get_strtime(time0):
                    table_father.display(self, "接待人意见：接待人意见日期与举报时间栏不一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '接待人意见', '接待人意见日期与举报时间栏不一致')

        text = self.contract_tables_content["承办部门负责人意见"]
        if text.strip() == "" or text.strip() == "/":
            table_father.display(self, "承办部门负责人意见：”举报记录表_承办部门负责人意见“不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r负责人意见', '”承办部门负责人意见“不能为空')
        else:
            pattern = re.compile(r'(.*)签名.*')
            yijian = re.findall(pattern, text)
            if yijian == [''] or yijian == [] or yijian[0].strip() == "":
                table_father.display(self, "承办部门负责人意见：”承办部门负责人意见“不能为空", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r负责人意见', '”承办部门负责人意见“不能为空')
            sign, date = tyh.sign_date(text)
            if sign == [''] or sign == [] or sign[0].strip() == "":
                table_father.display(self, "承办部门负责人意见：承办部门负责人未签名", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r负责人意见', '承办部门负责人未签名')
            if date == [""] or date == [] or date[0].strip() == "":
                table_father.display(self, "承办部门负责人意见：承办部门负责人未填写日期", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r负责人意见', '承办部门负责人未填写日期')
            else:
                date = date[0]
                if self.get_strtime(date) != self.get_strtime(time0):
                    table_father.display(self, "承办部门负责人意见：“承办部门负责人”日期与举报时间栏不一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r负责人意见',
                                       "”承办部门负责人“日期（" + date + "）与举报时间（" + time0 + "）栏不一致")

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


if __name__ == '__main__':

    my_prefix = r"C:\Users\Zero\Desktop\烟草输入_文件夹\副本\\"
    list = os.listdir(my_prefix)
    if "举报记录表_.docx" in list:
        ioc = table1(my_prefix, my_prefix)
        contract_file_path = os.path.join(my_prefix, "举报记录表_.docx")
        ioc.check(contract_file_path, "举报记录表_.docx")
