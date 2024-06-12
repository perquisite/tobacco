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
    'nameTimePlaceRight': '1．“烟草专卖局”：不为空，填写承办单位名称；2．“年 月 日”：不为空，与立案报告表案发日期保持一致；3．“局在xx查”：不为空，与立案报告表案发地点保持一致',
    'placeRight': '1．“内到xx ”：不为空，与案件承办单位保持一致；2．“（地址xx 联系人：xx 联系电话：xx ）”：不为空，由审查人员主观审查',
    'formRight': '1．“（地址xx 联系人：xx 联系电话：xx ）”：不为空，由审查人员主观审查；2．“品种规格”：不为空，与先行登记保存通知书品种规格保持一致；3．“数量”：不为空，与先行登记保存通知书记载的数量保持一致；4．“共计：（品种）”：不为空，与先行登记保存通知书共计数量保持一致；5．“总计：（数量）”：不为空，与先行登记保存通知书共计数量保持一致',
    'timeRight': '“年 月 日”：不为空，不能早于案发时间，晚于调查终结时间，由审查人员主观审查',
}

class table13(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix
        

        self.all_to_check = [
            "self.nameTimePlaceRight()",
            "self.placeRight()",
            "self.formRight()",
            "self.timeRight()"
        ]

    

    def check(self, contract_file_path,file_name_real):
        print("正在审查公告，审查结果如下：")
        self.mw = win32com.client.Dispatch("Word.Application")
        self.doc = self.mw.Documents.Open(self.my_prefix +file_name_real)

        data = DocxData(file_path=contract_file_path)
        self.contract_text = data.text
        # print('段落文本内容：', self.contract_text)
        doc = docx.Document(self.my_prefix + file_name_real)
        all_cell = []
        for t in doc.tables:
            for row in t.rows:
                cells = []
                for cell in row.cells:
                    cells.append(
                        cell.text.replace("\n", "").replace("\t", "").replace("\r", "").replace(" ", ""))
                all_cell.append(cells)
        self.tabels_content = all_cell
        # print(self.contract_tables_content)
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
        print(file_name_real+"审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result

    def nameTimePlaceRight(self):
        pattern = r"(.*)烟草专卖局[\S\s\n]*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"表头：表头不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "烟草专卖局", '表头不能为空')
        else:
            table_father.display(self,"表头：表头不为空", "green")

        pattern = r"(.*日)，我局在*"
        time = re.findall(pattern, self.contract_text)[0]
        timex = re.findall(pattern, self.contract_text)[0]
        if tyh.get_strtime(time) == False:
            table_father.display(self,"正文：时间错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, timex, '时间错误')
        else:
            if tyh.startTime(self.source_prifix) == False:
                table_father.display(self,"正文：立案报告表案发时间无法获取", "red")
            else:
                time0 = tyh.startTime(self.source_prifix)
                out_put_date = time0
                time = tyh.get_strtime(time)
                out_put_time = time
                time = time.split('-')
                time0 = time0.split("-")
                if time[0] != time0[0] or time[1] != time0[1] or time[2] != time0[2]:
                    table_father.display(self,'正文：时间与立案报告表案发日期【'+str(out_put_time)+'】不一致', "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, timex, '时间与立案报告表案发日期【'+str(out_put_time)+'】不一致')
                else:
                    table_father.display(self,"正文：时间与立案报告表案发日期一致", "green")

        pattern = r".*我局在(.*)查获*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：案发地点不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "我局在", '案发地点不能为空')
        else:
            table_father.display(self,"正文：案发地点不为空", "green")
            textx = re.findall(pattern, self.contract_text)[0]
            if tyh.startPlace(self.source_prifix) == False:
                table_father.display(self,"正文：无法获取立案报表表案发地点", "red")
            else:
                if text[0].replace(" ", "") != tyh.startPlace(self.source_prifix):
                    table_father.display(self,"正文：地点与立案报告表地点不一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, textx, '地点与立案报告表地点【'+tyh.startPlace(self.source_prifix)+'】不一致')
                else:
                    table_father.display(self,"正文：地点与立案报告表地点一致", "green")

    def placeRight(self):
        pattern = r".*30日内到(.*)[（]地址：.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：地址不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "30日内到", '地址不能为空')
        else:
            table_father.display(self,"正文：地址不为空", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, "30日内到", '主观审查，请与立案件报告表中的 案件承办单位 一致')

        pattern = r".*地址：(.*)联系人：.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"正文：地址不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "地址：", '地址不能为空')
        else:
            table_father.display(self,"正文：地址不为空", "green")

        pattern = r".*联系人：(.*)联系电话：.*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"联系人：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "联系人：", '联系人不能为空')
        else:
            table_father.display(self,"联系人：不为空", "green")

        pattern = r".*联系电话：(.*)[）].*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"联系电话：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "联系电话：", '联系电话：不能为空')
        else:
            table_father.display(self,"联系电话：不为空", "green")

    def isNullRight(self, l):
        if (l[0].strip() == "" and l[1].strip() != "") or (l[1].strip() == "" and l[2].strip() != "") \
                or (l[2].strip() == "" and l[3].strip() != "") or (l[3].strip() == "" and l[2].strip() != ""):
            return False
        else:
            return True

    def formRight(self):
        form = self.tabels_content[1:-1]
        dict = {}
        sum = 0.0
        all = 0.0
        for i in range(0, len(form)):
            if self.isNullRight(form[i]) == False:
                table_father.display(self,"表格：表格错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "品种规格", '表格错误')
                return
            else:
                if form[i][0].strip() != "":
                    dict[form[i][0]] = float(form[i][1].replace("条",""))
                    sum += 1.0
                    all += float(form[i][1].replace("条",""))
                if form[i][2].strip() != "":
                    dict[form[i][2]] = float(form[i][3].replace("条",""))
                    sum += 1.0
                    all += float(form[i][3].replace("条",""))

        # if os.path.exists(self.source_prifix + "证据先行登记保存批准书_.docx") == 1:
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

            if dict1 != dict:
                table_father.display(self,"表格：品种规格或数量与先行登记保存通知书不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "品种规格", '品种规格或数量与先行登记保存通知书不一致')
            else:
                table_father.display(self,"表格：品种规格或数量与先行登记保存通知书一致", "green")
            # print(dict)
            # print(dict1)

        sum0 = 0.0
        all0 = 0.0
        text0 = self.tabels_content[-1][0]
        pattern = r".*共计：(.*)个[（]品种[）]*"
        text = re.findall(pattern, text0)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"表格：共计：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "共计", '共计：不能为空')
        else:
            text=tyh.ch2num(text[0].replace(" ", ""))
            if text==None:
                table_father.display(self,"表格：共计：（数量）数无法识别", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "共计", '共计：（数量）数无法识别')
            else:
                sum0 = float(text)

        pattern = r".*总计：(.*)条[（]数量[）]"
        text = re.findall(pattern, text0)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"表格：总计：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "总计", '总计：不能为空')
        else:
            text=tyh.ch2num(text[0].replace(" ", ""))
            if text==None:
                table_father.display(self,"表格：总计：（数量）数无法识别", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "总计", '总计：（数量）数无法识别')
            else:
                all0 = float(text)

        if sum!=sum0:
            table_father.display(self,"表格：共计：错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "共计：", '共计：错误')
        else:
            table_father.display(self,"表格：共计：正确", "green")

        if all!=all0:
            table_father.display(self,"表格：总计：错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "总计", '总计：错误')
        else:
            table_father.display(self,"表格：总计：正确", "green")

    def timeRight(self):
        pattern = r".*[\S\s\n](.*年.*月.*日*)"
        timex = re.findall(pattern, self.contract_text)[1]
        time = re.findall(pattern, self.contract_text)[1].replace(" ","")
        time = chinese_to_date(time)
        if time == None:
            table_father.display(self,"结尾时间：结尾时间格式错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, timex, '结尾时间格式错误')
        else:
            table_father.display(self,"结尾时间：结尾时间格式正确", "green")





if __name__ == '__main__':

    my_prefix = "C:/Users/Zero/Desktop/烟草文书demo/2021184117_崇烟立2021第1号/"
    list = os.listdir(my_prefix)
    my_prefix = "C:\\Users/Zero/Desktop/烟草文书demo/2021184117_崇烟立2021第1号/"
    if "公告_.docx" in list:
        ioc = table13(my_prefix, my_prefix)
        contract_file_path = os.path.join(my_prefix , "公告_.docx")
        ioc.check(contract_file_path,"公告_.docx")
