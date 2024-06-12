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
    'headRight': '表头不能为空',
    'basicRight': '1．“案由”：不为空， 应与案卷中《行政处罚决定书》、《案件处理审批表》、《延期（分期）缴纳罚款审批表》、《案件集体讨论记录》、《陈述申辩记录》、《结案报告》、《封面》中“案由一致”。' 
                  '2．“案发时间”：不为空，与《立案报告表》、《不予立案报告表》中“案发时间”一致'
                  '3．“案件当事人”： 不为空，与《卷宗封面》、《立案报告表》、《延长立案期限审批表》、《证据先行登记保存通知书》、《抽样取证物品清单》、《涉案烟草专卖品核价表》、《卷烟鉴别检验样品留样、损耗费用审批表》、《调查总结报告》、《延长案件调查终结审批表》、《案件处理审批表》、《先行登记保存证据处理通知书》、《行政处罚事先告知书》、《听证告知书》、《听证通知书》、《不予受理听证通知书》、《听证笔录》、《听证报告》、《当场行政处罚决定书》、《行政处罚决定书》、《违法物品销毁记录表》、《加处罚款决定书》、《延期（分期）缴纳罚款审批表》、《结案报告》、《封面》中“当事人”一致'
                  '4．“案件文号”： 不为空，填写卷宗“立案报告表”中领导批准立案的日期。 与《涉案烟草专卖品核价表》、《调查总结报告》、《延长案件调查终结审批表》、《案件处理审批表》、《卷宗封面》'
                  '5．“质检报告文号”：不为空，与《检验报告》编号一致',
    'formRight': '1．“规格及品名”：不为空，与《抽烟取证物品清单》中“品种规格”一致'
                 '2．“查获数量(条)”：不为空，与《抽烟取证物品清单》中“样品基数”一致'
                 '3．“留样、损耗数量（条）”：不为空，与《抽烟取证物品清单》中“抽样数量”-《检验报告》“退样”的差一致'
                 '4．“留样、损耗单价(元/条)”：不为空，应与《价格目录表》中对应卷烟批发价格一致'
                 '5．“留样、损耗金额（元）”：不为空，为本表中数量*单价的积'
                 '6．“合计”：不为空，对表中对应项求和'
                 '7．“合计金额”：不为空，与上面金额的合计一致',
    'opinionsOfTheUndertaker': '“承办人意见”：不为空，明确损耗金额为多少，具体值与标中合“计金额”一致；两个承办人签字，并注明日期；日期在处罚决定书《送达回证》签收日期之后',
    'opinionsOfTheDepartment': '“承办部门意见”：不为空，是否同意承办人意见，本部分责任人签字，并注明日期；加盖部门章；日期在处罚决定书《送达回证》签收日期之后',
    'opinionsOfTheF': '“财务部门审核意见”：不为空，是否同意，财务部门责任人签字，并注明日期；加盖部门章；日期在处罚决定书《送达回证》签收日期之后',
    'opinionsOfTheM': '“负责人审批”：不为空，是否同意，单位责任人签字，并注明日期；加盖单位章；日期在处罚决定书《送达回证》签收日期之后',
    'sign': '“收款人签名”：不为空，收款人签字，并注明日期；按捺手印，收款人与其他文书中当事人一致，或者与委托书中受托人一致；日期在处罚决定书《送达回证》签收日期之后',
}

class table14(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix
        self.all_money_ch=[]
        self.all_money=[]

        self.all_to_check = [
            'self.headRight()',
            'self.basicRight()',
            'self.formRight()',
            'self.opinionsOfTheUndertaker()',
            'self.opinionsOfTheDepartment()',
            'self.opinionsOfTheF()',
            'self.opinionsOfTheM()',
            'self.sign()'

        ]

    def check(self, contract_file_path,file_name_real):
        print('正在审查'+file_name_real+'，审查结果如下：')
        self.mw = win32com.client.Dispatch('Word.Application')
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
        print(file_name_real+'审查完毕\n')
        info_list_result = table_father.get_info_list(self)
        return info_list_result

    def headRight(self):
        text = self.contract_text
        pattern = r'.*[\r\S\s]№：(.*)'
        text = re.findall(pattern, text)
        if text == [] or text[0].strip() == '' or text[0].replace(' ', '') == '/':
            table_father.display(self, '表头：表头不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '№：', '表头不能为空')
        else:
            table_father.display(self, '表头：表头不为空', 'green')

    def basicRight(self):
        form = self.contract_tables_content
        reason = form['案由']
        time = form['案发时间']
        people = form['案件当事人']
        num1 = form['案件文号']
        num2 = form['质检报告文号']

        reason0 = ''
        time0 = ''
        people0 = ''
        date = ''

        # if os.path.exists(self.source_prifix + '立案报告表_.docx') == 1:
        if tyh.file_exists(self.source_prifix, '立案报告表'):
            data = tyh.file_exists_open(self.source_prifix, '立案报告表', DocxData)
            people0 = data.tabels_content['当事人']
            time0 = data.tabels_content['案发时间']
            reason0 = data.tabels_content['案由']
            text = data.tabels_content['负责人意见']

            sign, date = tyh.sign_date(text)
            if date == [''] or date == [] or date[0].strip() == '':
                table_father.display(self, '负责人意见未注明日期', 'red')
            else:
                date = date[0].replace(' ', '')

        if reason.strip() != '' and reason.strip() != '/':
            if reason != reason0:
                table_father.display(self, '案由：案由与立案报告表等文书中记载的案由不一致', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '案 由', '案由与立案报告表等文书中记载的案由【'+str(reason0)+'】不一致')
            else:
                table_father.display(self, '案由：案由与立案报告表等文书中记载的案由一致', 'green')
        else:
            table_father.display(self, '案由：案由不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '案 由', '案由不能为空')

        if time.strip() != '' and time.strip() != '/':
            out_put_time = time0
            time0 = time0.replace('年', '-').replace('月', '-').replace('日', '-').replace('时', '-').replace('分',
                                                                                                          ' ').replace(
                '/', '-').strip().replace(
                ' ', '')
            time0 = time0.split('-')
            for i in range(0, len(time0)):
                time0[i] = int(time0[i])

            time = time.replace('年', '-').replace('月', '-').replace('日', ' ').replace(
                '/', '-').strip().replace(
                ' ', '')
            time = time.split('-')
            for i in range(0, len(time)):
                time[i] = int(time[i])

            if time != time0[0:3]:
                table_father.display(self, '案发时间：案发时间与立案报告表等文书中记载的案发时间不一致', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '案发时间', '案发时间与立案报告表等文书中记载的案发时间【'+str(out_put_time)+'】不一致')
            else:
                table_father.display(self, '案发时间：案发时间与立案报告表等文书中记载的案发时间一致', 'green')
        else:
            table_father.display(self, '案发时间：案发时间不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '案发时间', '案发时间不能为空')

        if people.strip() != '' and people.strip() != '/':
            if people != people0:
                table_father.display(self, '案件当事人：案件当事人与立案报告表等文书中记载的案件当事人【'+str(people0)+'】不一致', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '案  件\r当事人', '案件当事人与立案报告表等文书中记载的案件当事人【'+str(people0)+'】不一致')
            else:
                table_father.display(self, '案件当事人：案件当事人与立案报告表等文书中记载的案件当事人一致', 'green')
        else:
            table_father.display(self, '案件当事人：案件当事人不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '案  件\r当事人', '案件当事人不能为空')

        if num1.strip() != '' and num1.strip() != '/':
            if num1 != date:
                table_father.display(self, '案件文号：案件文号与立案报告表等文书中记载的案件文号【'+str(date)+'】不一致', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '案件文号', '案件文号与立案报告表等文书中记载的案件文号【'+str(date)+'】不一致')
            else:
                table_father.display(self, '案件文号：案件文号与立案报告表等文书中记载的案件文号一致', 'green')
        else:
            table_father.display(self, '案件文号：案件文号不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '案件文号', '案件文号不能为空')

        if num2.strip() == '' or num2.strip() == '/':
            table_father.display(self, '案件文号：质检报告文号不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '质检报告', '质检报告文号不能为空')
        else:
            table_father.display(self, '案件文号：质检报告文号不为空', 'green')
            tyh.addRemarkInDoc(self.mw, self.doc, '质检报告', '注意与《检验报告》编号一致')

    def formRight(self):
        form = self.contract_tables_content['卷烟质检样品损耗明细']
        num1 = 0.0
        num2 = 0.0
        num3 = 0.0
        num4 = 0.0
        dict = {}
        dictx = {}
        for f in form[:-2]:
            if self.isNullRight(f) == False:
                table_father.display(self, '表格：表格错误', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '卷烟质检样品损耗明细', '表格错误')
                return
            else:
                if f[0] != '':
                    dict[f[0]] = float(f[1])
                    dictx[f[0]] = float(f[3])
                    num1 += float(f[1])
                    num2 += float(f[2])
                    num3 += float(f[3])
                    num4 += float(f[4])

        # if os.path.exists(self.source_prifix + '先行登记保存批准书_.docx') == 1:

        tyh.addRemarkInDoc(self.mw, self.doc, '留样、损耗', '未得到检验报告，主管审查留样、损耗数量是否与《抽烟取证物品清单》中“抽样数量”-《检验报告》“退样”的差一致')

        if tyh.file_exists(self.source_prifix, '先行登记保存批准书'):
            doc = tyh.file_exists_open(self.source_prifix, '先行登记保存批准书', docx.Document)
            all_cell = []
            for t in doc.tables:
                for row in t.rows:
                    cells = []
                    for cell in row.cells:
                        cells.append(
                            cell.text.replace('\n', '').replace('\t', '').replace('\r', '').replace(' ', ''))
                    all_cell.append(cells)
            form0 = all_cell

            dict1 = {}
            form0 = form0[1:]
            i = 0
            while '共计' not in form0[i][0]:
                if form0[i][0].strip() != '':
                    dict1[form0[i][0]] = float(form0[i][2])
                if form0[i][3].strip() != '':
                    dict1[form0[i][3]] = float(form0[i][5])
                i += 1
            if dict1 != dict:
                table_father.display(self, '表格：规格及品名 或 查获数量 与证据先行登记保存批准书不一致', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '规格及品名', '规格及品名 或 查获数量 与证据先行登记保存批准书不一致')

        # if os.path.exists(self.source_prifix + '涉案烟草专卖品核价表_.docx') == 1:
        if tyh.file_exists(self.source_prifix, '涉案烟草专卖品核价表'):
            doc = tyh.file_exists_open(self.source_prifix, '涉案烟草专卖品核价表', docx.Document)

            all_cell = []
            for t in doc.tables:
                for row in t.rows:
                    cells = []
                    for cell in row.cells:
                        cells.append(
                            cell.text.replace('\n', '').replace('\t', '').replace('\r', '').replace(' ', ''))
                    all_cell.append(cells)
            form = all_cell
            form = form[1:-1]
            # print(form)
            for i in range(0, len(form)):
                if self.isNullRight(form[i]) == False:
                    table_father.display(self, '表格：表格格式错误', 'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, '序号', '表格格式错误')
                    return
            dicty = {}
            for i in range(0, len(form)):
                if form[i][0].strip() != '':
                    dicty[form[i][1]] = float(form[i][3])

            if dicty != dictx:
                table_father.display(self, '表格：对应价格错误', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '单价(元/条)', '对应价格错误')
            else:
                table_father.display(self, '表格：对应价格正确', 'green')

        form = self.contract_tables_content['卷烟质检样品损耗明细'][-2]
        flag = 0
        if float(num1) != float(form[1]):
            flag = 1
            table_father.display(self, '表格：查获数量(条) 合计错误', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '合    计', '查获数量(条) 合计错误')
        if float(num2) != float(form[2]):
            flag = 1
            table_father.display(self, '表格：数量（条） 合计错误', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '合    计', '数量（条） 合计错误')
        if float(num3) != round(float(form[3]), 1):
            flag = 1
            table_father.display(self, '表格：单价(元/条) 合计错误', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '合    计', '单价(元/条) 合计错误')
        if float(num4) != round(float(form[4]), 1):
            flag = 1
            table_father.display(self, '表格：金额（元） 合计错误', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '合    计', '金额（元）合计错误')
        if flag:
            table_father.display(self, '表格：合计正确', 'green')

        pattern = r'.*[（]￥(.*)[）]'
        text = self.contract_tables_content['卷烟质检样品损耗明细'][-1][1]
        self.all_money = re.findall(pattern, text)

        if self.all_money == [] or self.all_money[0].strip() == '' or self.all_money[0].replace(' ', '') == '/':
            self.all_money = []
            table_father.display(self, '表格：合计金额不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '￥', '合计金额不能为空')
        else:
            self.all_money = float(self.all_money[0])
            if self.all_money != num4:
                table_father.display(self, '表格：合计金额错误', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '￥', '合计金额错误')

            pattern = r'.*[（]大写[）](.*)[（]'
            text = self.contract_tables_content['卷烟质检样品损耗明细'][-1][1]
            self.all_money_ch = re.findall(pattern, text)
            if self.all_money_ch == [] or self.all_money_ch[0].strip() == '' or self.all_money_ch[0].replace(' ', '') == '/':
                self.all_money_ch = []
                table_father.display(self, '表格：合计金额(大写)不能为空', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '（大写）', '合计金额（大写）不能为空')
            else:
                self.all_money_ch=self.all_money_ch[0]
                if self.all_money_ch != tyh.formatCurrency(self.all_money):
                    table_father.display(self, '表格：合计金额(大写)错误', 'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, '（大写）', '合计金额(大写)错误')


    def opinionsOfTheUndertaker(self):
        text0 = self.contract_tables_content['承办人意见']
        pattern = '(.*)签名*'
        opinion = re.findall(pattern, text0)
        if opinion == [] or opinion[0].strip() == '' or opinion[0].replace(' ', '') == '/':
            table_father.display(self, '承办人意见：承办人意见不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '承办人意见', '承办人意见不能为空')
            return
        else:
            text = opinion[0]
            if self.all_money == []:
                table_father.display(self, '承办人意见：合计金额为空，请检查', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '承办人意见', '合计金额为空，请检查')
            elif self.all_money_ch in text or str(self.all_money) in text:
                table_father.display(self, '承办人意见：明确损耗金额', 'green')
            else:
                table_father.display(self, '承办人意见：未明确损耗金额或者损耗金额合计错误', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '承办人意见', '未明确损耗金额或者损耗金额合计错误')

            pattern = r'.*签名：(.*)年.*'
            text = re.findall(pattern, text0)
            if text == [] or text[0].strip() == '' or text[0].replace(' ', '') == '/' or text[0].strip().rstrip(
                    string.digits) == '':
                table_father.display(self, '承办人意见：签名：不能为空', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '承办人意见', '签名：不能为空')
            else:
                table_father.display(self, '承办人意见：签名：不为空', 'green')
                tyh.addRemarkInDoc(self.mw, self.doc, '承办人意见', '主观审查是否两人签名')

            pattern = r'.*签名：(.*)'
            text = re.findall(pattern, text0)[0]
            text = text.replace('年', '-').replace('月', '-').replace('日', ' ').replace('/', '-').strip()
            text = tyh.subChar(text)
            date1 = tyh.get_strtime_with_(text)
            if date1 == False:
                table_father.display(self, '承办人意见：承办人意见时间错误', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '承办人意见', '承办人意见时间错误')
                return
            # if os.path.exists(self.source_prifix + '送达回证_.docx') == 1:
            if tyh.file_exists(self.source_prifix, '送达回证'):
                data = tyh.file_exists_open(self.source_prifix, '送达回证', DocxData)
                time = data.tabels_content['签收日期']
                out_put_date = time
                if time.strip() == '':
                    table_father.display(self, '承办人意见：送达回证_签收日期为空', 'red')
                    return
                elif tyh.get_strtime(time.strip()) == False:
                    table_father.display(self, '承办人意见：送达回证_签收日期错误', 'red')
                else:
                    time = tyh.get_strtime(time.strip()).split('-')
                    time = 365 * int(time[0]) + 31 * int(time[1]) + int(time[2])
                date1 = tyh.get_strtime_with_(date1.strip()).split('-')
                date1 = 365 * int(date1[0]) + 31 * int(date1[1]) + int(date1[2])
                if date1 < time:
                    table_father.display(self, '承办人意见：日期应在处罚决定书《送达回证》签收日期【'+str(out_put_date)+'】之后', 'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, '承办人意见', '日期应在处罚决定书《送达回证》签收日期【'+str(out_put_date)+'】之后')
                else:
                    table_father.display(self, '承办人意见：日期在处罚决定书《送达回证》签收日期之后', 'green')

    def opinionsOfTheDepartment(self):
        text0 = self.contract_tables_content['承办部门意见']
        pattern = '(.*)签名*'
        opinion = re.findall(pattern, text0)
        if opinion == [] or opinion[0].strip() == '' or opinion[0].replace(' ', '') == '/':
            table_father.display(self, '承办部门意见：承办部门意见意见不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门意见', '承办部门意见意见不能为空')
            return
        else:
            pattern = r'.*签名：(.*)年.*'
            text = re.findall(pattern, text0)
            if text == [] or text[0].strip() == '' or text[0].replace(' ', '') == '/' or text[0].strip().rstrip(
                    string.digits) == '':
                table_father.display(self, '承办部门意见签名：不能为空', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门意见', '签名：不能为空')
            else:
                table_father.display(self, '承办部门意见签名：不为空', 'green')

            pattern = r'.*签名：(.*)'
            text = re.findall(pattern, text0)[0]
            text = text.replace('年', '-').replace('月', '-').replace('日', ' ').replace('/', '-').strip()
            text = tyh.subChar(text)
            date1 = tyh.get_strtime_with_(text)
            if date1 == False:
                table_father.display(self, '承办部门意见：承办部门意见意见时间错误', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门意见', '承办部门意见意见时间错误')
                return
            # if os.path.exists(self.source_prifix + '送达回证_.docx') == 1:
            if tyh.file_exists(self.source_prifix, '送达回证'):
                data = tyh.file_exists_open(self.source_prifix, '送达回证', DocxData)

                time = data.tabels_content['签收日期']
                out_put_date = time
                if time.strip() == '':
                    table_father.display(self, '承办部门意见：送达回证_签收日期为空', 'red')
                    return
                elif tyh.get_strtime(time.strip()) == False:
                    table_father.display(self, '承办部门意见：送达回证_签收日期错误', 'red')
                else:
                    time = tyh.get_strtime(time.strip()).split('-')
                    time = 365 * int(time[0]) + 31 * int(time[1]) + int(time[2])
                date1 = tyh.get_strtime_with_(date1.strip()).split('-')
                date1 = 365 * int(date1[0]) + 31 * int(date1[1]) + int(date1[2])
                if date1 < time:
                    table_father.display(self, '承办部门意见：日期应在处罚决定书《送达回证》签收日期【'+str(out_put_date)+'】之后', 'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, '承办部门意见', '日期应在处罚决定书《送达回证》签收日期【'+str(out_put_date)+'】之后')
                else:
                    table_father.display(self, '承办部门意见：日期在处罚决定书《送达回证》签收日期之后', 'green')

    def opinionsOfTheF(self):
        text0 = self.contract_tables_content['财务部门审核意见']
        pattern = '(.*)签名*'
        opinion = re.findall(pattern, text0)
        if opinion == [] or opinion[0].strip() == '' or opinion[0].replace(' ', '') == '/':
            table_father.display(self, '财务部门审核意见：财务部门审核意见不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '财务部门审核\r意见', '财务部门审核意见不能为空')
            return
        else:
            pattern = r'.*签名：(.*)年.*'
            text = re.findall(pattern, text0)
            if text == [] or text[0].strip() == '' or text[0].replace(' ', '') == '/' or text[0].strip().rstrip(
                    string.digits) == '':
                table_father.display(self, '财务部门审核意见签名：不能为空', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '财务部门审核\r意见', '签名：不能为空')
            else:
                table_father.display(self, '财务部门审核意见签名：不为空', 'green')

            pattern = r'.*签名：(.*)'
            text = re.findall(pattern, text0)[0]
            text = text.replace('年', '-').replace('月', '-').replace('日', ' ').replace('/', '-').strip()
            text = tyh.subChar(text)
            date1 = tyh.get_strtime_with_(text)
            if date1 == False:
                table_father.display(self, '财务部门审核意见：财务部门审核意见时间错误', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '财务部门审核\r意见', '财务部门审核意见时间错误')
                return
            # if os.path.exists(self.source_prifix + '送达回证_.docx') == 1:
            if tyh.file_exists(self.source_prifix, '送达回证'):
                data = tyh.file_exists_open(self.source_prifix, '送达回证', DocxData)

                time = data.tabels_content['签收日期']
                out_put_date = time
                if time.strip() == '':
                    table_father.display(self, '财务部门审核意见：送达回证_签收日期为空', 'red')
                    return
                elif tyh.get_strtime(time.strip()) == False:
                    table_father.display(self, '财务部门审核意见：送达回证_签收日期错误', 'red')
                else:
                    time = tyh.get_strtime(time.strip()).split('-')
                    time = 365 * int(time[0]) + 31 * int(time[1]) + int(time[2])
                date1 = tyh.get_strtime_with_(date1.strip()).split('-')
                date1 = 365 * int(date1[0]) + 31 * int(date1[1]) + int(date1[2])
                if date1 < time:
                    table_father.display(self, '财务部门审核意见：日期应在处罚决定书《送达回证》签收日期之后', 'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, '财务部门审核\r意见', '日期应在处罚决定书《送达回证》签收日期【'+str(out_put_date)+'】之后')
                else:
                    table_father.display(self, '财务部门审核意见：日期在处罚决定书《送达回证》签收日期【'+str(out_put_date)+'】之后', 'green')

    def opinionsOfTheM(self):
        text0 = self.contract_tables_content['负责人审批']
        pattern = '(.*)签名*'
        opinion = re.findall(pattern, text0)
        if opinion == [] or opinion[0].strip() == '' or opinion[0].replace(' ', '') == '/':
            table_father.display(self, '负责人审批：负责人审批不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人\r审批', '负责人审批不能为空')
            return
        else:
            pattern = r'.*签名：(.*)年.*'
            text = re.findall(pattern, text0)
            if text == [] or text[0].strip() == '' or text[0].replace(' ', '') == '/' or text[0].strip().rstrip(
                    string.digits) == '':
                table_father.display(self, '负责人审批签名：不能为空', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '负责人\r审批', '签名：不能为空')
            else:
                table_father.display(self, '负责人审批签名：不为空', 'green')

            pattern = r'.*签名：(.*)'
            text = re.findall(pattern, text0)[0]
            text = text.replace('年', '-').replace('月', '-').replace('日', ' ').replace('/', '-').strip()
            text = tyh.subChar(text)
            date1 = tyh.get_strtime_with_(text)
            if date1 == False:
                table_father.display(self, '负责人审批：负责人审批时间错误', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '负责人\r审批', '负责人审批时间错误')
                return
            # if os.path.exists(self.source_prifix + '送达回证_.docx') == 1:
            if tyh.file_exists(self.source_prifix, '送达回证'):
                data = tyh.file_exists_open(self.source_prifix, '送达回证', DocxData)
                time = data.tabels_content['签收日期']
                out_put_date = time
                if time.strip() == '':
                    table_father.display(self, '负责人审批：送达回证_签收日期为空', 'red')
                    return
                elif tyh.get_strtime(time.strip()) == False:
                    table_father.display(self, '负责人审批：送达回证_签收日期错误', 'red')
                else:
                    time = tyh.get_strtime(time.strip()).split('-')
                    time = 365 * int(time[0]) + 31 * int(time[1]) + int(time[2])
                date1 = tyh.get_strtime_with_(date1.strip()).split('-')
                date1 = 365 * int(date1[0]) + 31 * int(date1[1]) + int(date1[2])
                if date1 < time:
                    table_father.display(self, '负责人审批：日期应在处罚决定书《送达回证》签收日期【'+str(out_put_date)+'】之后', 'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, '负责人\r审批', '日期应在处罚决定书《送达回证》签收日期【'+str(out_put_date)+'】之后')
                else:
                    table_father.display(self, '负责人审批：日期在处罚决定书《送达回证》签收日期之后', 'green')

    def sign(self):
        text0 = self.contract_tables_content['收款人签名']
        pattern = r'.*签名：(.*)年.*'
        text = re.findall(pattern, text0)
        sign = ''
        if text == [] or text[0].strip() == '' or text[0].replace(' ', '') == '/' or text[0].strip().rstrip(
                string.digits) == '':
            table_father.display(self, '收款人签名签名：不能为空', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '收款人签名', '签名：不能为空')
        else:
            sign = text[0].strip().rstrip(string.digits)
            table_father.display(self, '收款人签名签名：不为空', 'green')

        pattern = r'.*签名：(.*)'
        text = re.findall(pattern, text0)[0]
        text = text.replace('年', '-').replace('月', '-').replace('日', ' ').replace('/', '-').strip()
        text = tyh.subChar(text)
        date1 = tyh.get_strtime_with_(text)
        if date1 == False:
            table_father.display(self, '收款人签名：收款人签名时间错误', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '收款人签名', '收款人签名时间错误')
            return
        # if os.path.exists(self.source_prifix + '送达回证_.docx') == 1:
        if tyh.file_exists(self.source_prifix, '送达回证'):
            data = tyh.file_exists_open(self.source_prifix, '送达回证', DocxData)

            time = data.tabels_content['签收日期']
            out_put_date=time
            if time.strip() == '':
                table_father.display(self, '收款人签名：送达回证_签收日期为空', 'red')
                return
            elif tyh.get_strtime(time.strip()) == False:
                table_father.display(self, '收款人签名：送达回证_签收日期错误', 'red')
            else:
                time = tyh.get_strtime(time.strip()).split('-')
                time = 365 * int(time[0]) + 31 * int(time[1]) + int(time[2])
            date1 = tyh.get_strtime_with_(date1.strip()).split('-')
            date1 = 365 * int(date1[0]) + 31 * int(date1[1]) + int(date1[2])
            if date1 < time:
                table_father.display(self, '收款人签名：日期应在处罚决定书《送达回证》签收日期【'+str(out_put_date)+'】之后', 'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '收款人签名', '日期应在处罚决定书《送达回证》签收日期【'+str(out_put_date)+'】之后')
            else:
                table_father.display(self, '收款人签名：日期在处罚决定书《送达回证》签收日期之后', 'green')

        # if os.path.exists(self.source_prifix + '立案报告表_.docx') == 1:
        if tyh.file_exists(self.source_prifix, '立案报告表'):
            data = tyh.file_exists_open(self.source_prifix, '立案报告表', DocxData)

            sign1 = data.tabels_content['当事人']
            if sign1.strip() == '':
                table_father.display(self, '收款人签名：立案报告表_当事人为空', 'red')
                return
            else:
                if sign1 != sign:
                    table_father.display(self, '收款人签名：收款人签名不是当事人', 'red')
                    tyh.addRemarkInDoc(self.mw, self.doc, '收款人签名', '收款人签名不是当事人')
                else:
                    table_father.display(self, '收款人签名：收款人签名是当事人', 'green')

    def isNullRight(self, l):
        if l[0].strip() == '':
            for i in range(0, len(l)):
                if l[i].strip() != '':
                    return False
            return True
        else:
            for i in range(0, len(l)):
                if l[i].strip() == '':
                    return False
            return True


if __name__ == '__main__':
    my_prefix = 'C:\\Users\\Zero\\Desktop\\副本\\'
    list = os.listdir(my_prefix)
    if '卷烟鉴别检验样品留样、损耗费用审批表_.docx' in list:
        ioc = table14(my_prefix, my_prefix)
        contract_file_path = my_prefix + '卷烟鉴别检验样品留样、损耗费用审批表_.docx'
        ioc.check(contract_file_path)
