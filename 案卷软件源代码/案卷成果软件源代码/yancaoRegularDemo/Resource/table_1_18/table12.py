import string

import win32com.client

from yancaoRegularDemo.Resource.ReadFile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.OCR_IDCard import OCR_IDCard
from yancaoRegularDemo.Resource.tools.TimeOperator import TimeOper
from yancaoRegularDemo.Resource.tools.utils import *
from warnings import simplefilter

simplefilter(action='ignore', category=FutureWarning)

function_description_dict = {
    'checkIDCard': '对比身份证信息与《结案报告表》中信息',
    'formRight': '1．“（证据粘贴处）”：不为空，由审查人员主观审查'
                 '2．“说明事项”：由审查人员主观审查'
                 '3．“复制（提取）地点 ”：不为空，与对应文书中的证据一致'
                 '4．“复制（提取）时间：  年    月    日    时    分”：不为空，应与检查（勘验）检查日期保持一致'
                 '5．“执法人员及执法证号”：不为空，应与检查（勘验）执法人员姓名、执法证号保持一致',
    'nameRight': '“烟草专卖局”：不为空，填写承办单位名称',
    'elseRight': '“年 月 日”：不为空，不能早于案发时间，晚于结案时间，由审查人员主观审查',
}

class table12(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.source_prefix = source_prefix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix
        self.file_name_real = ""

        self.all_to_check = [
            "self.checkIDCard()",
            "self.nameRight()",
            "self.formRight()",
            "self.elseRight()"
        ]

    def check(self, contract_file_path, file_name_real):
        self.file_name_real = file_name_real
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

    def checkIDCard(self):
        #print(self.contract_tables_content)
        for i in self.contract_tables_content:
            for j in i:
                if "身份证" in j:
                    self.goCheckID()
                    break

    def goCheckID(self):
        # 1、先检查 身份证姓名和住址 与《结案报告表》中的当事人和住址保持一致
        # 先提取《结案报告表》中的当事人姓名
        file_name = '结案报告表'
        if tyh.file_exists(self.source_prefix, file_name):
            data = tyh.file_exists_open(self.source_prefix, file_name, DocxData)
            name_paper = data.tabels_content['当事人']
            address_paper = data.tabels_content['地址']
            #print('1、' + name_paper + ' ' + address_paper)
            pic_path = self.source_prefix + "picture/" + self.file_name_real.strip(".docx").strip(".doc") + "/word/media/"
            #print("图片文件夹 " + pic_path)
            pic_name_list = []
            for root, dirs, files in os.walk(pic_path):
                pic_name_list = files
            #print(pic_path + pic_name_list[0], pic_path + pic_name_list[1])
            id_ocr = OCR_IDCard(pic_path + pic_name_list[0], pic_path + pic_name_list[1])
            id_name = id_ocr.getName()
            id_address = id_ocr.getAddress()
            id_date = id_ocr.getExpiringDate()
            #print('2、' + id_name + ' ' + id_address)
            # 判断姓名、地址是否一致
            if id_name == name_paper:
                pass
            else:
                table_father.display(self, "身份证图片提取的姓名：从身份证图片提取的姓名（"+id_name+"）与案卷的“当事人”"+name_paper+"不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '', '从身份证图片提取出姓名【'+id_name+'】与案卷的“当事人”【'+str(name_paper)+'】不一致')
            if id_address == address_paper:
                pass
            else:
                table_father.display(self, "身份证图片提取的姓名：从身份证图片提取的地址（"+id_address+"）与案卷的“地址”一栏（"+address_paper+"）不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '', '从身份证图片提取的地址【'+id_address+'】与案卷的“地址”【'+str(address_paper)+'】不一致')
        # 2、再检查 身份证到期时间应在提取时间之后
            # print(id_date)
            # 转化日期表达格式
            str_list = list(id_date)
            str_list.insert(4, '-')
            str_list.insert(7, '-')
            id_date = ''.join(str_list)
            t = TimeOper()
            if t.time_order(id_date, t.getLocalDate()) >= 0:
                pass
            else:
                table_father.display(self, "身份证：身份证已过期！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '', '身份证已过期！')



    def nameRight(self):
        pattern = r'(.*)烟草专卖局*'
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self, "烟草专卖局：烟草专卖局不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "烟草专卖局", '烟草专卖局不能为空')
        else:
            table_father.display(self, "烟草专卖局：烟草专卖局不为空", "green")

    def formRight(self):
        tyh.addRemarkInDoc(self.mw, self.doc, "证据复制(提取)单", '证据粘贴处主观审查')
        tyh.addRemarkInDoc(self.mw, self.doc, "说明事项：", '主观审查')

        pattern = r".*复制（提取）地点：(.*)复制（提取）时间.*"
        place = re.findall(pattern, self.contract_tables_content[3][0])
        if place == [] or place[0].strip() == "" or place[0].replace(" ", "") == '/':
            table_father.display(self, "复制（提取）地点：复制（提取）地点：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "复制（提取）地点：", '复制（提取）地点：不能为空')
            place = None
        else:
            place = place[0]
            table_father.display(self, "复制（提取）地点：复制（提取）地点：不为空", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, "复制（提取）地点：", '主观审查对应文书中的证据是否一致')

        pattern = r"复制（提取）时间：(.*)"
        time = re.findall(pattern, self.contract_tables_content[3][0])
        if time == [] or time[0].strip() == "" or time[0].replace(" ", "") == '/':
            table_father.display(self, "复制（提取）时间：复制（提取）时间：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "复制（提取）时间：", '复制（提取）时间：不能为空')
        else:
            if tyh.get_strtime_5(time[0]) == False:
                table_father.display(self, "复制（提取）时间：复制（提取）时间：格式错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "复制（提取）时间：", '复制（提取）时间：格式错误')
            else:
                # if os.path.exists(self.source_prefix + "检查（勘验）笔录_.docx") == 1:
                if tyh.file_exists(self.source_prefix, "检查（勘验）笔录"):
                    data = tyh.file_exists_open(self.source_prefix, "检查（勘验）笔录", DocxData)
                    text = data.text
                    pattern = r".*检查（勘验）时间：(.*)至.*"
                    time1 = re.findall(pattern, text)
                    pattern1 = r".*检查（勘验）时间：.*至(.*)"
                    time2 = re.findall(pattern1, text)
                    if time1 == [] or time1[0].strip() == "" or time1[0].replace(" ", "") == '/' or time2 == [] or \
                            time2[0].strip() == "" or time2[0].replace(" ", "") == '/':
                        table_father.display(self, "复制（提取）时间：不能提取检查（勘验）时间", "red")
                    else:
                        if tyh.get_strtime_5(time1[0]) == False or tyh.get_strtime_5(time2[0]) == False:
                            table_father.display(self, "复制（提取）时间：检查（勘验）时间：格式错误", "red")
                        else:
                            time1 = tyh.get_strtime_5(time1[0])
                            time2 = tyh.get_strtime_5(time2[0])
                            time = tyh.get_strtime_5(time[0])
                            if time and time1 and time2:
                                if tyh.time_differ_5(time1, time) < 0 or tyh.time_differ_5(time, time2) < 0:
                                    table_father.display(self, "复制（提取）时间：复制（提取）时间应与检查（勘验）检查日期保持一致", "red")
                                    tyh.addRemarkInDoc(self.mw, self.doc, "复制（提取）时间：", '复制（提取）时间应与检查（勘验）检查日期不一致')
                                else:
                                    table_father.display(self, "复制（提取）时间：复制（提取）时间与检查（勘验）检查日期保持一致", "green")
                            else:
                                table_father.display(self, "复制（提取）时间：复制（提取）时间应或检查（勘验）检查日期提取错误", "red")
                                tyh.addRemarkInDoc(self.mw, self.doc, "复制（提取）时间：", '复制（提取）时间应或检查（勘验）检查日期提取错误')

        text = self.contract_tables_content[2][0]
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self, "说明事项：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, text, '说明事项：不能为空')
        else:
            if place not in text:
                table_father.display(self, "复制（提取）地点：复制（提取）地点没有与对应文书中的证据一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "复制（提取）地点：", '复制（提取）地点没有与对应文书中的证据一致')

        pattern = r".*执法人员及执法证号：(.*)"
        text = re.findall(pattern, self.contract_tables_content[4][0])
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self, "执法人员及执法证号：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "执法人员及执法证号：", '执法人员及执法证号：不能为空')
        else:
            table_father.display(self, "执法人员及执法证号：不为空", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, "执法人员及执法证号：", '与检查（勘验）执法人员姓名、执法证号保持一致')

    def elseRight(self):
        # tyh.addRemarkInDoc(self.mw, self.doc, "印章：", '不为空，加盖案件承办单位公章')
        pattern = r".*[\S\s\n](.*年.*月.*日*)"
        timex = re.findall(pattern, self.contract_text)[0]
        time = re.findall(pattern, self.contract_text)[0].replace(" ", "")
        time = chinese_to_date(time)
        if time == None:
            table_father.display(self, "结尾时间：结尾时间格式错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, timex, '结尾时间格式错误')
        else:
            table_father.display(self, "结尾时间：结尾时间格式正确", "green")
            if tyh.startTime(self.source_prefix) == False or tyh.endTime(self.source_prefix) == False:
                table_father.display(self, "结尾时间：时间无法提取", "red")
            else:
                out_startTime = tyh.startTime(self.source_prefix)
                out_endTime = tyh.endTime(self.source_prefix)
                out_time = time
                startTime = tyh.startTime(self.source_prefix).split("-")
                endTime = tyh.endTime(self.source_prefix).split("-")
                time = time.split("-")

                # print(time)
                # print(startTime)
                # print(endTime)

                startTime = 1000 * int(startTime[0]) + 31 * int(startTime[1]) + int(startTime[2])
                endTime = 1000 * int(endTime[0]) + 31 * int(endTime[1]) + int(endTime[2])
                time = 1000 * int(time[0]) + 31 * int(time[1]) + int(time[2])

                if startTime <= time <= endTime:
                    table_father.display(self, "结尾时间：结尾时间晚于案发时间，早于结案时间", "green")
                    tyh.addRemarkInDoc(self.mw, self.doc, timex, '主观审查')
                else:
                    table_father.display(self, "结尾时间：结尾时间不能早于案发时间，晚于结案时间", "res")
                    tyh.addRemarkInDoc(self.mw, self.doc, timex, '结尾时间【'+str(out_time)+'】不能早于案发时间【'+str(out_startTime)+'】，晚于结案时间【'+str(out_endTime)+'】')


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    file = '证据复制提取单_'
    if file + ".docx" in list:
        ioc = table12(my_prefix, my_prefix)
        contract_file_path = my_prefix + file + ".docx"
        ioc.check(contract_file_path, file + ".docx")
