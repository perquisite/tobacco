import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
from yancaoRegularDemo.Resource.tools.utils import chinese_to_date
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh

function_description_dict = {
    'check_varieties_and_specifications': '涉案卷烟品种、规格、数量与《涉案烟草专卖品核价表》保持一致',
    'check_clause': '违法条款、处罚条款与《案件调查终结报告》保持一致。',
    'check_date': '作出日期与《听证告知书》作出日期保持一致或之后，“之后”指《听证告知书》作出日期3天后。',
}


# 行政处罚事先告知书
class Table25(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prifix = source_prifix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.mw.Visible = 0
        self.mw.DisplayAlerts = 0
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.check_varieties_and_specifications()",
            "self.check_clause()",
            "self.check_date()"
        ]

    def check_varieties_and_specifications(self):
        """
        作用：涉案卷烟品种、规格、数量与《涉案烟草专卖品核价表》保持一致
        """
        if not tyh.file_exists(self.source_prifix, "涉案烟草专卖品核价表"):
            table_father.display(self, "文件缺失：《涉案烟草专卖品核价表》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "涉案烟草专卖品核价表", DocxData)
            other_tabels_content = other_info.tabels_content

            # 首先判断物品的品种数量是否一致
            other_count = len(other_tabels_content["涉案烟草专卖品核价表-序号"])  # 《涉案烟草专卖品核价表》物品种类
            item_count_pattern = re.compile(r'计(.*)个品种')
            item_count = re.findall(item_count_pattern, self.contract_text)[0]
            if int(item_count) != other_count:
                table_father.display(self, "涉案卷烟品种数目：与《涉案烟草专卖品核价表》不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '个品种',
                                   '涉案卷烟品种数目与《涉案烟草专卖品核价表》不一致,《涉案烟草专卖品核价表》中物品数量为：' + str(
                                       item_count))
            else:
                # 物品的品种数量一致,判断物品品种是否一致
                i = 0
                exit_label = True
                for item_name in other_tabels_content["涉案烟草专卖品核价表-品种规格"]:
                    item_quantity = other_tabels_content["涉案烟草专卖品核价表-数量(条)"][i]
                    i += 1
                    item_for_search = item_name + item_quantity
                    if item_for_search not in self.contract_text:
                        exit_label = False
                        # table_father.display(self, "× 《行政处罚事先告知书》的品种种类、规格与《涉案烟草专卖品核价表》不一致", "red")
                        # tyh.addRemarkInDoc(self.mw, self.doc, '个品种', '品种种类、规格与《涉案烟草专卖品核价表》不一致')
                        table_father.display(self,
                                             "品种规格：与《涉案烟草专卖品核价表》不一致，未在此表中查询到的关键词为：" + str(
                                                 item_for_search),
                                             "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, '个品种',
                                           '品种种类或数目与《涉案烟草专卖品核价表》不一致，未在此表中查询到的关键词为：' + item_for_search)
                        break
                if exit_label:
                    table_father.display(self, "《行政处罚事先告知书》的品种种类或数目：正确。与《涉案烟草专卖品核价表》一致",
                                         "green")

    def check_clause(self):
        """
        作用：违法条款、处罚条款与《案件调查终结报告》保持一致。
            a“违法条款”识别：“违反XXX”。
            b“处罚条款”识别：“依据XXX进行如下行政处罚”。
        """
        if not tyh.file_exists(self.source_prifix, "案件调查终结报告"):
            table_father.display(self, "文件缺失：《案件调查终结报告》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "案件调查终结报告", DocxData)
            other_tabels_content = other_info.tabels_content
            illegal_pattern = re.compile(r'违反了(.*?)规定')
            punish_pattern = re.compile(r'依据(.*?)规定')

            if re.findall(illegal_pattern, self.contract_text) is not []:
                illegal_text = re.findall(illegal_pattern, self.contract_text)[0]
                if illegal_text not in other_tabels_content['案件性质']:
                    tyh.addRemarkInDoc(self.mw, self.doc, '违反了', '违法条款与《案件调查终结报告》不一致')
            else:
                tyh.addRemarkInDoc(self.mw, self.doc, '行政处罚事先告知书',
                                   '未识别到“违法条款”，识别模板为【违反了xxx规定】')

            if re.findall(punish_pattern, self.contract_text) is not []:
                punish_text = re.findall(punish_pattern, self.contract_text)[0]
                if punish_text not in other_tabels_content['处罚依据']:
                    tyh.addRemarkInDoc(self.mw, self.doc, '依据', '处罚依据与《案件调查终结报告》不一致')
            else:
                tyh.addRemarkInDoc(self.mw, self.doc, '行政处罚事先告知书',
                                   '未识别到“处罚条款”，识别模板为【依据xxx规定】')

    def check_date(self):
        """
        作用：作出日期与《听证告知书》作出日期保持一致或之后，“之后”指《听证告知书》作出日期3天后。
        """
        if not tyh.file_exists(self.source_prifix, "听证告知书"):
            table_father.display(self, "文件缺失：《听证告知书》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "听证告知书", DocxData)
            other_tabels_text = other_info.text
            time_pattern = re.compile(r'.*(二.{3}年.{1,2}月.{1,3}日).*')
            this_file_time = re.findall(time_pattern, self.contract_text)
            other_file_time = re.findall(time_pattern, other_tabels_text)
            if this_file_time == [] or this_file_time is None:
                table_father.display(self, "作出日期：未找到作出日期，作出日期请采用中文格式的年月日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '行政处罚事先告知书',
                                   '未找到作出日期，作出日期请采用中文格式的年月日')
            elif other_file_time == [] or other_file_time is None:
                table_father.display(self, "作出日期：《听证告知书》未找到作出日期，作出日期请采用中文格式的年月日", "red")
            else:
                this_file_time_date = chinese_to_date(this_file_time[0])
                other_file_time_date = chinese_to_date(other_file_time[0])
                time_differ = tyh.time_differ(this_file_time_date, other_file_time_date)
                if time_differ == 0:
                    table_father.display(self, "作出日期：正确。与《听证告知书》作出日期保持一致", "green")
                elif time_differ >= 3:
                    table_father.display(self,
                                         "作出日期：正确。在《听证告知书》作出日期之后（“之后”指《听证告知书》作出日期3天后。",
                                         "green")
                else:
                    table_father.display(self,
                                         "作出日期：未与《听证告知书》作出日期保持一致或之后（“之后”指《听证告知书》作出日期3天后。",
                                         "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, this_file_time,
                                       '作出日期错误，未与《听证告知书》作出日期保持一致或之后（“之后”指《听证告知书》作出日期3天后。《听证告知书》作出日期为：' + other_file_time_date)

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
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
        # self.mw.Quit()
        print("《行政处罚事先告知书》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    if "行政处罚事先告知书_.docx" in list:
        ioc = Table25(my_prefix, my_prefix)
        contract_file_path = my_prefix + "行政处罚事先告知书_.docx"
        ioc.check(contract_file_path, '行政处罚事先告知书_.docx')
