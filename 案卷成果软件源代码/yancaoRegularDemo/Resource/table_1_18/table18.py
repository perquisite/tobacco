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
    'headRight': '文书编号不为空。',
    'nameRight': '当事人名字/名称一般与《立案报告表》中当事人名字/名称一致，若不一致，预警提示。',
    'textRight': '1、请求哪个单位协助主观审查；2、案由一般与《立案报告表》中案由一致，若不一致，预警提示；3、调查内容不为空，需要调查的内容必须具体明确，主观判断。；4、联系人、联系电话不为空，且与《公告》中联系人、联系电话 一致。',
    'timeRight': '印章不为空。时间在《立案报告》中“案发时间”之后。',
}

class table18(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix
        

        self.all_to_check = [
            "self.headRight()",
            "self.nameRight()",
            "self.textRight()",
            "self.timeRight()"

        ]

    

    def check(self, contract_file_path,file_name_real):
        print("正在审查"+file_name_real+"，审查结果如下：")
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
        #self.mw.Quit()
        print(file_name_real+"审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result

    def headRight(self):
        pattern = r".*烟协[〔](.*)[〕]第.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"表头：表头不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "烟协", '表头不能为空')
        else:
            table_father.display(self,"烟协不为空", "green")

        pattern = r".*第(.*)号.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"表头：第几号不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "第", '第几号不能为空')
        else:
            table_father.display(self,"表头：第几号不为空", "green")

    def nameRight(self):
        pattern = ".*号[\s\S\n](.*)：[\s\S\n]本局.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"表头：名字不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "：", '名字不能为空')
        else:
            table_father.display(self,"表头：名字不为空", "green")
            # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                doc = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
                form = doc.tabels_content
                self.dangshiren = form["当事人"]
                self.anyou = form["案由"]
                if text[0].replace(" ", "") != self.dangshiren.replace(" ", ""):
                    table_father.display(self,'表头：名字与立案报告表当事人【'+str(self.dangshiren)+'】不同', "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, text[0], '名字与立案报告表当事人【'+str(self.dangshiren)+'】不同')
                else:
                    table_father.display(self,"表头：名字与立案报告表当事人相同", "green")

    def textRight(self):
        pattern = r".*本局在调查处理(.*)一案中.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：案由不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "调查处理", '案由不能为空')
        else:
            textx = text[0]
            text = text[0].replace(" ", "")
            if text != self.anyou:
                table_father.display(self,'正文：应与《立案报告表》中“案由【'+str(self.anyou)+'】”一致', "red")
                tyh.addRemarkInDoc(self.mw, self.doc, textx, '应与《立案报告表》中“案由【'+str(self.anyou)+'】”一致')
            else:
                table_father.display(self,"正文：与《立案报告表》中“案由 ”一致", "green")

        pattern = r".*需要调查(.*)。请你单位予以协助。*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：需要调查什么不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "需要调查", '需要调查什么不能为空')
        else:
            table_father.display(self,"正文：需要调查什么不为空", "green")

        people = ""
        phone = ""

        pattern = r".*联系人：(.*)联系电话*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"联系人：联系人不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "联系人", '联系人不能为空')
        else:
            table_father.display(self,"联系人：联系人不为空", "green")
            people = text[0].replace(" ","")

        pattern = r".*联系电话：(.*)"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"联系电话：联系电话不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "联系电话", '联系电话不能为空')
        else:
            table_father.display(self,"联系电话：联系电话不为空", "green")
            phone = text[0].replace(" ","")

        # if os.path.exists(self.source_prifix + "公告_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "公告"):
            doc = tyh.file_exists_open(self.source_prifix, "公告", DocxData)
            text0 = doc.text
            pattern = r".*联系人：(.*)联系电话*"
            text = re.findall(pattern, text0)
            if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
                table_father.display(self,"联系人：公告联系人不能为空", "red")
            else:
                if text[0].replace(" ","") != people:
                    table_father.display(self,'联系人：联系人与公告中【'+str(text[0].replace(" ",""))+'】不同', "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "联系人", '联系人与公告中【'+str(text[0].replace(" ",""))+'】不同')
                else:
                    table_father.display(self,"联系人：联系人与公告相同", "green")

            pattern = r".*联系电话：(.*)[）].*"
            text = re.findall(pattern, text0)
            if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
                table_father.display(self,"联系电话：公告联系电话不能为空", "red")
            else:
                if text[0].replace(" ","") != phone:
                    table_father.display(self,'联系电话：联系电话与公告中【'+str(text[0].replace(" ",""))+'】不同', "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "联系电话", '联系电话与公告中【'+str(text[0].replace(" ",""))+'】不同')
                else:
                    table_father.display(self,"联系电话：联系电话与公告相同", "green")

    def timeRight(self):
        pattern = r".*[\S\s\n](.*年.*月.*日*)"
        time=re.findall(pattern,self.contract_text)[0]
        timex = time
        time=tyh.changeDate(time)
        # print(time)
        if tyh.get_strtime_with_(time)==False:
            table_father.display(self,"时间：时间错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, time, '时间错误')
        else:
            time=tyh.get_strtime_with_(time).split("-")
            time=int(time[0])*365+int(time[1])*31+int(time[2])

            stime=tyh.startTime(self.source_prifix)

            if tyh.get_strtime_5_with_(stime)==False:
                table_father.display(self,"时间：立案报告表案发时间错误", "red")
            else:
                out_put_date = stime
                stime=stime.split("-")
                stime=int(stime[0])*365+int(stime[1])*31+int(stime[2])
                if stime>=time:
                    table_father.display(self,'时间：时间在《立案报告》中“案发时间【'+str(out_put_date)+'】”之前', "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, timex, '时间在《立案报告》中“案发时间【'+str(out_put_date)+'】”之前')
                else:
                    if stime < time:
                        table_father.display(self,"时间：时间在《立案报告》中“案发时间”之后", "green")


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Zero\\Desktop\\副本\\"
    list = os.listdir(my_prefix)
    if "协助调查函_.docx" in list:
        ioc = table18(my_prefix, my_prefix)
        contract_file_path = my_prefix + "协助调查函_.docx"
        ioc.check(contract_file_path)
