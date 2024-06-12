import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import chinese_to_date

function_description_dict = {
    'check_serial_number': '文书编号与《证据先行登记保存通知书》保持一致。',
    'check_result_of_handling': '写明处理的结果：物品的品种、规格和数量与《涉案烟草专卖品核价表》保持一致',
    'check_date': '作出日期在《行政处罚事先告知》作出之日3天后。',
    'check_sign_time': '当事人签名时间、执法人员签名时间与本文书作出时间一致。'
}


# 先行登记保存证据处理通知书
class Table23(table_father):
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
            "self.check_serial_number()",
            "self.check_result_of_handling()",
            "self.check_date()",
            "self.check_sign_time()"
        ]

    def check_serial_number(self):
        """
        作用：“  烟存通[ ]第   号”与《证据先行登记保存通知书》文书编号保持一致。
        """
        # 未找到《证据先行登记保存通知书》
        if not tyh.file_exists(self.source_prifix, "证据先行登记保存通知书"):
            table_father.display(self, "文件缺失：《证据先行登记保存通知书》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "证据先行登记保存通知书", DocxData)
            other_tabels_content = other_info.tabels_content
            other_text = other_info.text

    def check_result_of_handling(self):
        """
        作用：写明处理的结果：物品的品种、规格和数量与《涉案烟草专卖品核价表》保持一致
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
                table_father.display(self, "物品品种数目：与《涉案烟草专卖品核价表》不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '个品种',
                                   '物品品种数量与《涉案烟草专卖品核价表》不一致,《涉案烟草专卖品核价表》中物品数量为：' + item_count)
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
                        table_father.display(self,
                                             "品种规格：与《涉案烟草专卖品核价表》不一致，未在此表中查询到的关键词为：" + item_for_search,
                                             "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, '做出如下处理：',
                                           '品种种类或数目与《涉案烟草专卖品核价表》不一致，未在此表中查询到的关键词为：' + item_for_search)
                        break
                if exit_label:
                    table_father.display(self,
                                         "《先行登记保存证据处理通知书》品种种类或数目：正确。与《涉案烟草专卖品核价表》一致",
                                         "green")

    def check_date(self):
        """
        作用：作出日期在《行政处罚事先告知》作出之日3天后。
        """
        # 判断限期届满之日是否为法定节假日，是的话将该日期改为节假日之后的第一天
        if not tyh.file_exists(self.source_prifix, "行政处罚事先告知书"):
            table_father.display(self, "文件缺失：《行政处罚事先告知书》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "行政处罚事先告知书", DocxData)
            other_tabels_text = other_info.text
            other_time_parttern = re.compile(r'.*(二.{3}年.{1,2}月.{1,3}日).*')
            time_parttern = re.compile(r'.*(\d{4}年\d{1,2}月\d{1,2}日).*')
            this_file_time = re.findall(time_parttern, self.contract_text)[0]
            other_file_time = re.findall(other_time_parttern, other_tabels_text)[-1]
            if this_file_time == '' or this_file_time is None:
                table_father.display(self, "作出日期：未找到作出日期，作出日期请采用中文格式的年月日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '先行登记保存证据处理通知书',
                                   '未找到作出日期，作出日期请采用中文格式的年月日')
            elif other_file_time == '' or other_file_time is None:
                table_father.display(self, "作出日期：《行政处罚事先告知书》未找到作出日期，作出日期请采用x年x月x日格式",
                                     "red")
            else:
                this_file_time_date = tyh.get_strtime(this_file_time)
                other_file_time_date = chinese_to_date(other_file_time)
                time_differ = tyh.time_differ(this_file_time_date, other_file_time_date)
                if time_differ < 3:
                    table_father.display(self, "作出日期：应在《行政处罚事先告知》作出之日3天后", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, this_file_time,
                                       "作出日期应在《行政处罚事先告知》作出之日3天后,《行政处罚事先告知》作出之日为：" + other_file_time_date)
                else:
                    table_father.display(self, "作出日期：正确。在《行政处罚事先告知》作出之日3天后", "green")

    def check_sign_time(self):
        """
        作用：当事人签名时间、执法人员签名时间与本文书作出时间一致。
        """

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
        print("《先行登记保存证据处理通知书》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    if "先行登记保存证据处理通知书_.docx" in list:
        ioc = Table23(my_prefix, my_prefix)
        contract_file_path = my_prefix + "先行登记保存证据处理通知书_.docx"
        ioc.check(contract_file_path, "先行登记保存证据处理通知书_.docx")
