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
    'nameAndTime': '1、“烟草专卖局”：不为空，填写承办单位名称；2．“年 月 日”：不为空，与立案报告表案发日期保持一致；',
    'formRight': '1．“价格如下： ”：不为空，根据案发时间/卷烟品种（价格目录中匹配不到的）选择上半年、下半年、去年全年'
                 '2．“序号”：不为空，填写格式为阿拉伯数字“1,2,3,4…”'
                 '3．“品种规格”：不为空，与先行登记保存通知书品种规格保持一致'
                 '4．“数量”：不为空，与先行登记保存通知书记载的数量保持一致'
                 '5．“单价（元）”：不为空，与价格目录中对应品牌的批发价保持一致'
                 '6．“合计（元）”：不为空，应等于数量*单价'
                 '7．“合计”：不为空，等于各项相加总数',
    'peopleAndSignRight': '1．“经办人”：不为空，两人签名，与案件承办人保持一致'
                          '2．“当事人”：不为空，应与立案报告表等文书中记载的当事人名称保持一致，无当事人签字时注明理由'
                          '3．“（印章）”：不为空，加盖案件承办单位公章'
                          '4．“年 月 日”：不为空，该日期应大于或等于质检报告日期并小于案件调查终结报告中承办部门意见一栏中的日期',
}

class table11(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix
        

        self.all_to_check = [
            "self.nameAndTime()",
            "self.formRight()",
            "self.peopleAndSignRight()"

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

    def nameAndTime(self):
        pattern = r"(.*)烟草专卖局*"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"名字:名字不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "烟草专卖局于", '名字不能为空')
        else:
            table_father.display(self,"名字:名字不为空", "green")

        pattern = r".*于(.*)查获*"
        time = re.findall(pattern, self.contract_text)[0]
        if tyh.get_strtime(time) == False:
            table_father.display(self,"正文时间：时间错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "年", '时间错误')
        else:
            # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                data = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
                time0 = data.tabels_content["案发时间"]
                if time0 == "" or time0 == "/":
                    table_father.display(self,"正文时间：立案报告表 时间错误", "red")
                else:
                    time = time.replace("年", "-").replace("月", "-").replace("日", " ").replace(
                        "/", "-").strip().replace(
                        " ", "")
                    time = time.split("-")
                    for i in range(0, len(time)):
                        time[i] = int(time[i])
                    out_put_date = time0
                    time0 = time0.replace("年", "-").replace("月", "-").replace("日", "-").replace("时", "-").replace("分",
                                                                                                                  " ").replace(
                        "/", "-").strip().replace(
                        " ", "")
                    time0 = time0.split("-")
                    for i in range(0, len(time0)):
                        time0[i] = int(time0[i])
                    if time0[0:3] != time:
                        table_father.display(self,"正文时间：时间与立案时间不一致", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, "年", '时间与《立案报告表》立案时间【'+str(out_put_date)+'】不一致')
                    else:
                        table_father.display(self,"正文时间：时间与立案时间一致", "green")

    def formRight(self):
        doc = docx.Document(self.my_prefix + "涉案烟草专卖品核价表_.docx")
        all_cell = []
        for t in doc.tables:
            for row in t.rows:
                cells = []
                for cell in row.cells:
                    cells.append(
                        cell.text.replace("\n", "").replace("\t", "").replace("\r", "").replace(" ", ""))
                all_cell.append(cells)
        form = all_cell
        form = form[1:-1]
        # print(form)
        for i in range(0, len(form)):
            if self.isNullRight(form[i]) == False:
                table_father.display(self,"表格序号：表格格式错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "序号", '表格格式错误')
                return
        table_father.display(self,"表格序号：表格格式正确", "green")

        dict = {}
        num = 0.0
        all = 0.0

        for i in range(0, len(form)):
            if form[i][0].strip() != "":
                if form[i][0] != str(i + 1):
                    table_father.display(self,"表格序号：表格序号错误", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "序号", '表格序号错误')

                dict[form[i][1]] = float(form[i][2])

                if float(form[i][2]) * float(form[i][3]) != float(form[i][4]):
                    table_father.display(self,"表格：表格合计错误", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "合计", '表格合计错误')

                num += float(form[i][2])
                all += float(form[i][4])

        # if os.path.exists(self.source_prifix + "证据先行登记保存通知书_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "证据先行登记保存通知书"):
            doc = tyh.file_exists_open(self.source_prifix, "证据先行登记保存通知书", docx.Document)
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
            # print(dict)

            if dict != dict1:
                table_father.display(self,"表格：品种规格应或数量与先行登记保存通知书所记载的品种规格保持不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "品种规格", '品种规格应或数量与先行登记保存通知书所记载的品种规格保持不一致')
            else:
                table_father.display(self,"表格：品种规格应 数量与先行登记保存通知书所记载的品种规格保持一致", "green")

        num1 = float(self.contract_tables_content["涉案烟草专卖品核价表-全部数量合计"])
        all1 = float(self.contract_tables_content["涉案烟草专卖品核价表-全部金额合计"])

        # print(num)
        # print(all)

        if num != num1:
            table_father.display(self,"表格：合计总数量出错", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "数量", '合计总数量出错')

        if round(all,1) != all1:
            table_father.display(self,"表格：合计总金额出错", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "合计", '合计总金额出错')

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

    def peopleAndSignRight(self):
        pattern = r".*经办人：(.*)"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"经办人：经办人不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "经办人：", '经办人不能为空')
        else:
            tyh.addRemarkInDoc(self.mw, self.doc, "经办人：", '主观审查是否 两人签名，与案件承办人保持一致')

        pattern = r".*当事人：(.*)"
        text = re.findall(pattern, self.contract_text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            table_father.display(self,"当事人：当事人不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "当事人：", '当事人不能为空')
        else:
            # if os.path.exists(self.source_prifix + "立案报告表_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "立案报告表"):
                data = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)

                people = data.tabels_content["当事人"]
                if text[0] != people:
                    table_father.display(self,"当事人：当事人应与立案报告表等文书中记载的当事人名称保持一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "当事人：", '当事人应与立案报告表等文书中记载的当事人名称【'+str(people)+'】保持一致')

        pattern = r".*[\S\s\n](.*年.*月.*日*)"
        time = re.findall(pattern, self.contract_text)[1]
        timex = time
        time = time.replace('\r', '').replace('\n', '').replace('\t', '').replace(" ", "")
        time = chinese_to_date(time)
        if time == None:
            table_father.display(self,"结尾日期：结尾日期格式错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "（印章）", '结尾日期格式错误')
        else:
            d_text = ""
            # if os.path.exists(self.source_prifix + "调查终结报告_.docx") == 1:
            if tyh.file_exists(self.source_prifix, "调查终结报告"):
                data1 = tyh.file_exists_open(self.source_prifix, "调查终结报告", DocxData)
                text0 = data1.tabels_content["处理意见"]
                if text0 != "":
                    d_text = text0
                _, d_date = tyh.sign_date(d_text)
                if d_date == [""] or d_date == [] or d_date[0].strip() == "":
                    table_father.display(self,"结尾日期：案件调查终结报告处理意见栏未注明日期", "red")
                else:
                    d_date = tyh.get_strtime(d_date[0])
                    date = tyh.get_strtime_with_(time)
                    # print(date)
                    # print(d_date)
                    if date and d_date:
                        if tyh.time_differ(date, d_date) > 0:
                            table_father.display(self,"结尾日期：结尾时间不得在案件调查终结报告处理意见栏所注明的日期之后", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, timex, '结尾时间不得在案件调查终结报告处理意见栏所注明的日【'+str(d_date)+'】之后')
                        else:
                            table_father.display(self,"结尾日期：结尾时间在案件调查终结报告处理意见栏所注明的日期之前", "green")
                    else:
                        table_father.display(self, "结尾日期：时间提取错误", "red")

                    if tyh.file_exists(self.source_prifix,"质检报告")==False:
                        table_father.display(self, "质检报告：不存在质检报告", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, timex, '不存在质检报告')



if __name__ == '__main__':
    my_prefix = "C:\\Users\\Zero\\Desktop\\副本\\"
    list = os.listdir(my_prefix)
    if "涉案烟草专卖品核价表_.docx" in list:
        ioc = table11(my_prefix, my_prefix)
        contract_file_path = my_prefix + "涉案烟草专卖品核价表_.docx"
        ioc.check(contract_file_path)
