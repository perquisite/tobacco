import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.TimeOperator import TimeOper
from yancaoRegularDemo.Resource.tools.simple_content import Simple_Content

# 行政处理决定书
from yancaoRegularDemo.Resource.tools.tanweijia_function import is_exist_cover
from yancaoRegularDemo.Resource.tools.utils import is_valid_date

function_description_dict = {
    'check_Case': '案件基本情况：应当与《案件处理审批表》中“案件事实”记载的一致。',
    'check_Notice': '公告的过程与结果 与模板对比。',
    'check_Decision': '# 处理依据与决定 与模板对比。',
    'check_Time': '作出行政处理决定的烟草专卖局的名称和日期。日期应在《案件处理审批表》时间之后。'
}

class Table37(table_father):
    def __init__(self, my_prefix, source_prefix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prefix = source_prefix  # 2021-08-07版本新增
        self.contract_text = None
        self.contract_tables_content = None
        self.mw = win32com.client.Dispatch("Word.Application")

        self.all_to_check = [
            "self.check_Case()",
            "self.check_Notice()",
            "self.check_Decision()",
            "self.check_Time()"
        ]

    # 案件基本情况：应当与《案件处理审批表》中“案件事实”记载的一致。
    def check_Case(self):
        if not tyh.file_exists(self.source_prefix, "行政处理决定书"):
            table_father.display(self, "× " + "《行政处理决定书》.docx不存在", "red")
        else:
            if tyh.file_exists(self.source_prefix, "案件处理审批表"):
                #print(self.contract_text)
                data = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
                case1 = data.tabels_content['案件事实'].strip()
                #print(case1)
                if "违法事实：" in case1:
                    case1.strip("违法事实：")
                elif "违法事实:" in case1:
                    case1.strip("违法事实:")
                if case1 not in self.contract_text:
                    table_father.display(self, "案件基本情况：案件基本情况与《案件处理审批表》中“案件事实”记载的不一致！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "", "该文书的案件基本情况与《案件处理审批表》中“案件事实”记载的不一致！")

    # 公告的过程与结果 与模板对比。
    def check_Notice(self):
        if tyh.file_exists(self.source_prefix, "行政处理决定书"):
            # count = 0
            content = self.contract_text
            if "公告" not in content:
                table_father.display(self, "公告的过程与结果：该文书中没有体现“公告的过程”！请人工核查。", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "", "该文书中没有体现“公告的过程”！请人工核查。")
        # 将《公告》和本文书的落款时间对比，看是否是当事人公告逾期没来
        if not tyh.file_exists(self.source_prefix, "公告"):
            table_father.display(self, "公告的过程与结果：《公告》文件不存在，无法对比落款时间以判断当事人是否在公告后逾期未到！请人工核查。", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, "", "《公告》文件不存在，无法对比落款时间以判断当事人是否在公告后逾期未到！请人工核查。")
        else:
            data = tyh.file_exists_open(self.source_prefix, "公告", DocxData)
            raw_time1_1 = data.text.split()[-1]
            raw_time1_2 = data.text.split()[-2]
            # print(raw_time1_1)
            # print(raw_time1_2)
            if "年" in raw_time1_1 and "月" in raw_time1_1 and "日" in raw_time1_1:
                time1 = tyh.changeDate(raw_time1_1) # 公告里的时间
                raw_time1 = raw_time1_1
            elif "年" in raw_time1_2 and "月" in raw_time1_2 and "日" in raw_time1_2:
                time1 = tyh.changeDate(raw_time1_2)
                raw_time1 = raw_time1_2
            # print(time1)
            raw_time2 = self.contract_text.split()[-2]
            # print(raw_time2)
            time2 = tyh.changeDate(raw_time2)  # 本文书的时间
            # print(time2)
            if tyh.time_differ(time2, time1) > 30:
                pass
            else:
                table_father.display(self, "公告的过程与结果：《公告》的日期为“"+ raw_time1 + "”，而本文书日期为“" + raw_time2 + "”，故不符合公告后30日逾期的规则,请人工核查！", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, raw_time2, "《公告》的日期为“"+ raw_time1 + "”，而本文书日期为“" + raw_time2 + "”，故不符合公告后30日逾期的规则,请人工核查！")



    # 处理依据与决定 与模板对比。
    def check_Decision(self):
        if tyh.file_exists(self.source_prefix, "行政处理决定书"):
            # 依据
            if "《烟草专卖行政处罚程序规定》第五十八条" not in self.contract_text:
                table_father.display(self, "处理依据与决定：该文书中没有体现“处理依据”！请人工核查。", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "", "该文书中没有体现“处理依据”！请人工核查。")
            # 决定
            if "没收" in self.contract_text:  # 没收 是违规的 直接报错
                table_father.display(self, "处理依据与决定：“没收”属于违规的“处理决定”！请人工核查。", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "没收", "“没收”属于违规的“处理决定”！请人工核查。")
            else:
                if "销毁" not in self.contract_text and "变卖" not in self.contract_text:
                    table_father.display(self, "处理依据与决定：该文书中没有体现“处理决定”！请人工核查。", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, "", "该文书中没有体现“处理决定”！请人工核查。")

    # 作出行政处理决定的烟草专卖局的名称和日期。日期应在《案件处理审批表》时间之后。
    def check_Time(self):
        if tyh.file_exists(self.source_prefix, "行政处理决定书"):
            if not tyh.file_exists(self.source_prefix, "案件处理审批表"):
                table_father.display(self, "表格缺失：" + "《案件处理审批表》.docx不存在", "red")
            else:
                temp = self.contract_text.split()
                name = temp[0]
                # print(name)
                time_raw = temp[-2]
                time = tyh.changeDate(time_raw)
                data = tyh.file_exists_open(self.source_prefix, "案件处理审批表", DocxData)
                content = data.tabels_content['负责人意见']
                temp = re.search("(\s*\S+\s*)年(\s*\S+\s*)月(\s*\S+\s*)日", content)
                time1 = temp[1] + "-" + temp[2] + "-" + temp[3]
                t = TimeOper()
                if t.time_order(time, time1) <= 0:
                    table_father.display(self, "作出行政处理决定日期：作出行政处理决定日期应在《案件处理审批表》时间（"+time1+"）之后！", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, time_raw, "该文书作出行政处理决定日期应在《案件处理审批表》时间（"+time1+"）之后！")

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
        #self.mw = win32com.client.Dispatch("Word.Application")
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
        self.doc.Close()
        self.mw.Quit()
        print("《行政处理决定书》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\twj\Desktop\\test\\"
    list = os.listdir(my_prefix)
    if "行政处理决定书_.docx" in list:
        ioc = Table37(my_prefix, my_prefix)
        contract_file_path = my_prefix + "行政处理决定书_.docx"
        ioc.check(contract_file_path, "行政处理决定书_.docx")
