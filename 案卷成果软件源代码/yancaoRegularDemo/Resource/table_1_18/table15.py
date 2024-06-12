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
    'headRight': '“烟移〔 〕第  号”：不为空',
    'nameRight': '“xx”： 不为空，与《卷宗封面》、《立案报告表》、《延长立案期限审批表》、《证据先行登记保存通知书》、《抽样取证物品清单》、《涉案烟草专卖品核价表》、《卷烟鉴别检验样品留样、损耗费用审批表》、《调查总结报告》、《延长案件调查终结审批表》、《案件处理审批表》、《先行登记保存证据处理通知书》、《行政处罚事先告知书》、《听证告知书》、《听证通知书》、《不予受理听证通知书》、《听证笔录》、《听证报告》、《当场行政处罚决定书》、《行政处罚决定书》、《违法物品销毁记录表》、《加处罚款决定书》、《延期（分期）缴纳罚款审批表》、《结案报告》中“当事人”一致',
    'textRight': '1．“xx一案”：不为空；“当事人+案由”'
                 '2．“发现xx”：不为空，该案涉嫌刑事犯罪'
                 '3．“根据 xx的规定”：不为空，《中华人民共和国行政处罚法》第二十二条'
                 '4．“材料  份  页，移送财物清单   页。”：不为空',
    'signAndDate': '1．“经办人”：不为空，与承办人一致'
                   '2．“批准人”：不为空，与单位负责人一致'
                   '3．“（印章）”：不为空，为单位章'
                   '4．“年 月 日”：不为空；表现为“ XX年XX月XX日” ，应在《调查总结报告》时间之后。',
}

class table15(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix
        

        self.dangshiren = ""
        self.anyou = ""

        self.all_to_check = [
            "self.headRight()",
            "self.nameRight()",
            "self.textRight()",
            "self.signAndDate()"
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
        pattern = r".*烟移[〔](.*)[〕]第.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"表头：表头不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "烟移", '表头不能为空')
        else:
            table_father.display(self,"表头：烟移不为空", "green")

        pattern = r".*第(.*)号.*"
        text = re.findall(pattern, self.contract_text)
        print(text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"表头：第几号不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "第", '第几号不能为空')
        else:
            table_father.display(self,"表头：第几号不为空", "green")

    def nameRight(self):
        pattern = ".*号[\s\S\n](.*)：[\s\S\n]本局.*"
        text = re.findall(pattern, self.contract_text)
        # print(text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"名字：名字不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "：", '名字不能为空')
        else:
            table_father.display(self,"名字：名字不为空", "green")
            # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                data = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
                form = data.tabels_content
                self.dangshiren = form["当事人"]
                self.anyou = form["案由"]
                if text[0].replace(" ", "") != self.dangshiren.replace(" ", ""):
                    table_father.display(self,"名字：名字与立案报告表当事人不同", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "：", '名字与立案报告表当事人【'+str(text[0].replace(" ", ""))+'】不同')
                else:
                    table_father.display(self,"名字：名字与立案报告表当事人相同", "green")

    def textRight(self):
        pattern = ".*本局对(.*)一案进行调查时.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "一案", '不能为空')
        else:
            table_father.display(self,"正文：不为空", "green")
            if self.dangshiren not in text[0] or self.anyou not in text[0]:
                table_father.display(self,"正文：案由或当事人未在其中", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, text[0], '案由或当事人未在其中')
            else:
                table_father.display(self,"正文：案由或当事人在其中", "green")

        pattern = ".*发现(.*)。根据.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：发现不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "发现", '发现不能为空')
        else:
            table_father.display(self,"正文：发现不为空", "green")
            if "该案涉嫌刑事犯罪" not in text[0]:
                table_father.display(self,"正文：该案涉嫌刑事犯罪未在发现中", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, text[0], '该案涉嫌刑事犯罪未在发现中')
            else:
                table_father.display(self,"正文：该案涉嫌刑事犯罪在发现中", "green")

        pattern = r".*根据(.*)的规定.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "根据", '不能为空')
        else:
            table_father.display(self,"正文：不为空", "green")
            if "《中华人民共和国行政处罚法》第二十二条" not in text[0]:
                table_father.display(self,"正文：《中华人民共和国行政处罚法》第二十二条 不在规定中", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, text[0], '《中华人民共和国行政处罚法》第二十二条 不在规定中')
            else:
                table_father.display(self,"正文：《中华人民共和国行政处罚法》第二十二条 在规定中", "green")

        pattern = r".*材料(.*)份.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：份不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "份", '份不能为空')
        else:
            table_father.display(self,"正文：份不为空", "green")

        pattern = r".*份(.*)页，移送*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：页不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "页", '页不能为空')
        else:
            table_father.display(self,"正文：页不为空", "green")

        pattern = r".*移送财物清单(.*)页.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：移送财物清单几页不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "清单", '移送财物清单几页不能为空')
        else:
            table_father.display(self,"正文：移送财物清单几页不为空", "green")

    def signAndDate(self):
        pattern = r".*经办人：(.*)"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"经办人：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "经办人：", '经办人：不能为空')
        else:
            table_father.display(self,"经办人：不为空", "green")

        pattern = r".*批准人：(.*)"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"批准人：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "批准人：", '批准人：不能为空')
        else:
            table_father.display(self,"批准人：不为空", "green")

        pattern = r".*[\S\s\n](.*年.*月.*日*).*"
        timex = re.findall(pattern, self.contract_text)[0]
        time = re.findall(pattern, self.contract_text)[0].replace(" ", "").replace("\t", "")
        time = tyh.changeDate(time)
        if time == None or tyh.get_strtime_with_(time)==False:
            table_father.display(self,"结尾时间：结尾时间格式错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, timex, '结尾时间格式错误')
        else:
            table_father.display(self,"结尾时间：结尾时间格式正确", "green")
            # if os.path.exists(self.source_prifix + "案件调查终结报告_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "案件调查终结报告"):
                data = tyh.file_exists_open(self.source_prifix, "案件调查终结报告", DocxData)

                text = data.tabels_content["处理意见"]
                _, date = tyh.sign_date(text)
                if date == [] or date[0].strip() == "" or date[0].replace(" ", "") == '/':
                    table_father.display(self,"结尾时间：案件调查终结报告_时间获取失败", "red")
                else:
                    if tyh.get_strtime(date[0]) == False:
                        table_father.display(self,"结尾时间：案件调查终结报告_结尾时间格式错误", "red")
                    else:
                        date=tyh.get_strtime(date[0])
                        out_put_date = date
                        date1=date.split("-")
                        date1 = 365 * int(date1[0]) + 31 * int(date1[1]) + int(date1[2])

                        time=time.split("-")
                        time = 365 * int(time[0]) + 31 * int(time[1]) + int(time[2])

                        if time<date1:
                            table_father.display(self,"结尾时间：结尾时间应在《调查终结报告》时间【"+str(out_put_date)+"】之后。", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, timex, '结尾时间应在《调查终结报告》时间【'+str(out_put_date)+'】之后。')
                        else:
                            table_father.display(self,"结尾时间：结尾时间在《调查终结报告》时间之后。", "green")

if __name__ == '__main__':
    my_prefix = "C:\\Users\\Zero\\Desktop\\副本\\"
    list = os.listdir(my_prefix)
    if "案件移送函_.docx" in list:
        ioc = table15(my_prefix, my_prefix)
        contract_file_path = my_prefix + "案件移送函_.docx"
        ioc.check(contract_file_path)
