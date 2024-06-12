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
    'timeRight': '接收时间不为空，且与《案件移送函》作出时间一致或之后。',
    'reasonRight': '接收单位意见不为空。',
    'signRight': '1.经办人不为空，应有接收单位的2名经办人签字。；2.接收单位印章不为空，加盖接收单位的印章。即签字和印章两者都有。时间在《案件移送函》之后。',
}

class table16(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix
        

        self.dangshiren = ""
        self.anyou = ""

        self.all_to_check = [
            "self.timeRight()",
            "self.reasonRight()",
            "self.signRight()"

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

    def timeRight(self):
        pattern = r".*接收时间：(.*)"
        text = re.findall(pattern, self.contract_text)[0]
        if tyh.get_strtime(text) == False:
            table_father.display(self,"接收时间：错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "接收时间：", '接收时间：错误')
        else:
            # if os.path.exists(self.source_prifix + "案件移送函_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "案件移送函"):
                try:
                    data = tyh.file_exists_open(self.source_prifix, "案件移送函", DocxData)
                    pattern = r"(.*年.*月.*日*)"
                    self.time0 = re.findall(pattern, data.text)[0].replace(" ", "")
                    self.time0 = tyh.changeDate(self.time0)
                    time0 = self.time0.split("-")
                    time0 = 365 * int(time0[0]) + 31 *int(time0[1]) + int(time0[2])

                    time = tyh.get_strtime(text)
                    time = time.split("-")
                    time = 365 * int(time[0]) + 31 * int(time[1]) + int(time[2])

                    if (time0 > time):
                        table_father.display(self,'接收时间：接收时间在《案件移送函》作出时间【'+str(time0)+'】之前', "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "接收时间", '接收时间在《案件移送函》作出时间【'+str(time0)+'】之前')
                    else:
                        table_father.display(self,"接收时间：接收时间在《案件移送函》作出时间一样或之后", "green")
                except:
                    table_father.display(self, '接收时间：案件移送函_时间提取错误', "red")
                    tyh.addRemarkInDoc(self.mw, self.doc,"接收时间", "接收时间：案件移送函_时间提取错误")

    def reasonRight(self):
        pattern = r".*接收单位意见：(.*)"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"接收单位意见：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "接收单位意见：", '接收单位意见：不能为空')
        else:
            table_father.display(self,"接收单位意见：接收单位意见不为空", "green")

    def signRight(self):
        pattern = r".*经办人：(.*)"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"经办人：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "经办人：", '经办人：不能为空')
        else:
            table_father.display(self,"经办人：不为空", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, "经办人：", '主观审查是否有两个经办人签字')

        pattern = r".*[\S\s\n](.*年.*月.*日*).*"
        timex = re.findall(pattern, self.contract_text)[1]
        time = re.findall(pattern, self.contract_text)[1].replace(" ","")
        time=tyh.changeDate(time)
        if tyh.get_strtime_with_(time)==False:
            table_father.display(self,"结尾时间：结尾时间错误", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, timex, '结尾时间错误')
        else:
            if tyh.file_exists(self.source_prifix , "案件移送函") == 1:
                try:
                    time0 = self.time0.split("-")
                    time0 = 365 * int(time0[0]) + 31 *int(time0[1]) + int(time0[2])

                    time = time.split("-")
                    time = 365 * int(time[0]) + 31 * int(time[1]) + int(time[2])

                    if (time0 > time):
                        table_father.display(self,'结尾时间：结尾时间错误在《案件移送函》作出时间【'+str(self.time0)+'】之前', "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, timex, '结尾时间错误在《案件移送函》作出时间【'+str(self.time0)+'】之前')
                    else:
                        table_father.display(self,"结尾时间：结尾时间错误在《案件移送函》作出时间一样或之后", "green")
                except:
                    table_father.display(self, '结尾时间：案件移送函_时间提取错误', "red")
                    tyh.addRemarkInDoc(self.mw, self.doc,timex, "结尾时间：案件移送函_时间提取错误")


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Zero\\OneDrive\\案卷\\tyh\\"
    list = os.listdir(my_prefix)
    if "案件移送回执_.docx" in list:
        ioc = table16(my_prefix, my_prefix)
        contract_file_path = my_prefix + "案件移送回执_.docx"
        ioc.check(contract_file_path,"案件移送回执_.docx")
