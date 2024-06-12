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
    'headRight': '移送函编号不为空。  ',
    'formRight': '涉案卷烟品种、规格、数量不为空且与《涉案烟草专卖品核价表》一致，若出现移送非涉案卷烟，主观审查。',
    'signAndDateRight': '1、移送单位（印章）、接送单位（印章）不为空且不能相同;2、移送人、接送人由2名人员签字; 3、移送财物时间不为空且与《案件移送函》一致；接收财物时间不为空且与《案件移送回执》时间一致。',
}

class table17(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix
        

        self.all_to_check = [
            "self.headRight()",
            "self.formRight()",
            "self.signAndDateRight()"

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
        pattern = r'.*移送函编号：(.*)'
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"移送函编号：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "移送函编号：", '移送函编号：不能为空')
        else:
            table_father.display(self,"移送函编号：不为空", "green")

    def formRight(self):
        doc = docx.Document(self.my_prefix + "移送财物清单_.docx")
        all_cell = []
        for t in doc.tables:
            for row in t.rows:
                cells = []
                for cell in row.cells:
                    cells.append(
                        cell.text.replace("\n", "").replace("\t", "").replace("\r", "").replace(" ", ""))
                all_cell.append(cells)
        tabels_content = all_cell[1:-1]
        dict = {}
        for t in tabels_content:
            if self.isNullRight(t) == False:
                table_father.display(self,"表格：表格错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "品种", '表格错误')
                return
            else:
                if t[0].strip() != "":
                    dict[t[0] + t[1]] = float(t[2])


        dict1 = {}
        # if os.path.exists(self.source_prifix + "涉案烟草专卖品核价表_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "涉案烟草专卖品核价表"):
            doc = tyh.file_exists_open(self.source_prifix, "涉案烟草专卖品核价表", docx.Document)
            all_cell = []
            for t in doc.tables:
                for row in t.rows:
                    cells = []
                    for cell in row.cells:
                        cells.append(
                            cell.text.replace("\n", "").replace("\t", "").replace("\r", "").replace(" ", ""))
                    all_cell.append(cells)
            tabels_content = all_cell[1:-1]
            for t in tabels_content:
                if self.isNullRight(t) == False:
                    table_father.display(self,"涉案烟草专卖品核价表错误", "red")
                    return
                else:
                    if t[0].strip() != "":
                        dict1[t[1]] = float(t[2])

            for d in dict:
                if d not in dict1:
                    table_father.display(self,"表格"+str(d) + "与涉案烟草专卖品核价表不匹配", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "品种", str(d) + '没有出现在涉案烟草专卖品核价表中,主观审查')
                else:
                    if dict[d] != dict1[d]:
                        table_father.display(self,"表格"+str(d) + "与涉案烟草专卖品核价表不匹配", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "数量", str(d) + '与涉案烟草专卖品核价表数量不匹配')
                    else:
                        table_father.display(self,"表格"+str(d)+"表格与涉案烟草专卖品核价表匹配", "green")

    def signAndDateRight(self):
        doc = docx.Document(self.my_prefix + "移送财物清单_.docx")
        all_cell = []
        for t in doc.tables:
            for row in t.rows:
                cells = []
                for cell in row.cells:
                    cells.append(
                        cell.text.replace("\n", "").replace("\t", "").replace("\r", ""))
                all_cell.append(cells)
        tabels_content = all_cell[-1]

        tyh.addRemarkInDoc(self.mw, self.doc, "移送单位（印章）：", '主观审查接收单位（印章）：不能与移送单位（印章）相同')
        tyh.addRemarkInDoc(self.mw, self.doc, "接收单位（印章）：", '主观审查接收单位（印章）：不能与移送单位（印章）相同')
        text1=tabels_content[0]

        text2=tabels_content[2]

        pattern=r".*移送人：(.*)年.*"
        text=re.findall(pattern,text1)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/' or text[0].strip().rstrip(
                string.digits) == "":
            table_father.display(self,"移送人：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "移送人：", '移送人：不能为空')
        else:
            table_father.display(self,"移送人：不为空", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, "移送人：", '主观审查移送人：是否有两个')

        pattern=r".*接收人：(.*)年.*"
        text=re.findall(pattern,text2)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/'or text[0].strip().rstrip(
                string.digits) == "":
            table_father.display(self,"接收人：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "接收人：", '接收人：不能为空')
        else:
            table_father.display(self,"接收人：不为空", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, "接收人：", '主观审查接收人：是否有两个')

        pattern = r".*移送人：(.*年.*月.*日*)"
        time1=re.findall(pattern,text1)[0]

        time1 = time1.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip().replace(" ","")
        time1 = tyh.subChar(time1)

        if tyh.get_strtime_with_(time1)==False:
            table_father.display(self,"移送人：时间错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "移送人", '移送人：时间错误')
        else:
            time1=tyh.get_strtime_with_(time1)
            # if os.path.exists(self.source_prifix + "案件移送函_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "案件移送函"):
                doc = tyh.file_exists_open(self.source_prifix, "案件移送函", DocxData)
                text=doc.text
                pattern = r".*[\S\s\n](.*年.*月.*日*)"

                timex=re.findall(pattern,text)[0].replace(" ","")
                timex=tyh.changeDate(timex)
                if tyh.get_strtime_with_(timex)==False:
                    table_father.display(self,"移送人：案件移送函时间错误", "red")
                else:

                    if time1!=timex:
                        table_father.display(self,'移送人：移送人时间与《案件移送函》时间【'+str(timex)+'】不一致', "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "移送人", '移送人时间与《案件移送函》时间【'+str(timex)+'】不一致')
                    else:
                        table_father.display(self,"移送人：移送人时间与《案件移送函》时间一致", "green")

        pattern = r".*接收人：(.*年.*月.*日*)"
        time2=re.findall(pattern,text2)[0]

        time2 = time2.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip().replace(" ","")
        time2 = tyh.subChar(time2)

        if tyh.get_strtime_with_(time2)==False:
            table_father.display(self,"接收人：时间错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "接收人", '接收人：时间错误')
        else:
            time2=tyh.get_strtime_with_(time2)
            # if os.path.exists(self.source_prifix + "案件移送回执_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "案件移送回执"):
                doc = tyh.file_exists_open(self.source_prifix, "案件移送回执", DocxData)

                text=doc.text
                pattern = r".*[\S\s\n](.*年.*月.*日*)"

                timex=re.findall(pattern,text)[1].replace(" ","")
                timex=tyh.changeDate(timex)
                if tyh.get_strtime_with_(timex)==False:
                    table_father.display(self,"接收人：案件移送回执时间错误", "red")
                else:

                    if time2!=timex:
                        table_father.display(self,'接收人：接收财物时间与《案件移送回执》时间【'+str(timex)+'】不一致', "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "接收人：", '接收财物时间与《案件移送回执》时间【'+str(timex)+'】不一致')
                    else:
                        table_father.display(self,"接收人：接收财物时间与《案件移送回执》时间一致", "green")




    def isNullRight(self, l):
        if l[0].strip() == "":
            for i in range(0, len(l)):
                if l[i].strip() != "":
                    return False
            return True
        else:
            for i in range(0, len(l)):
                if l[i].strip() == "":
                    return False
            return True


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Zero\\Desktop\\副本\\"
    list = os.listdir(my_prefix)
    if "移送财物清单_.docx" in list:
        ioc = table17(my_prefix, my_prefix)
        contract_file_path = my_prefix + "移送财物清单_.docx"
        ioc.check(contract_file_path)
