import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh

function_description_dict = {
    'check_varieties_and_specifications': '品种、规格与《涉案烟草专卖品核价表》保持一致。数量一栏可能涉及送检损耗，主观审查',
    'check_date': '返还时间与接收时间一致，且时间与《先行登记保存证据处理通知书》日期一致或之后。',
}


# 涉案物品返还清单
class Table24(table_father):
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
            "self.check_date()"
        ]

    def check_varieties_and_specifications(self):
        """
        作用：品种、规格与《涉案烟草专卖品核价表》保持一致。数量一栏可能涉及送检损耗，主观审查
        """
        if not tyh.file_exists(self.source_prifix, "涉案烟草专卖品核价表"):
            table_father.display(self, "文件缺失：《涉案烟草专卖品核价表》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "涉案烟草专卖品核价表", DocxData)
            other_tabels_content = other_info.tabels_content

            # 首先判断物品的品种数量是否一致
            other_count = len(other_tabels_content["涉案烟草专卖品核价表-序号"])  # 《涉案烟草专卖品核价表》物品种类
            item_count = len(self.contract_tables_content["涉案物品返还清单-品种"])
            if item_count != other_count:
                table_father.display(self, "物品数目：与《涉案烟草专卖品核价表》中数量" + str(item_count) + "不一致",
                                     "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '品种',
                                   '物品数目与《涉案烟草专卖品核价表》不一致,《涉案烟草专卖品核价表》中物品数量为：' + str(
                                       item_count))
            else:
                # 物品的品种数量一致,判断物品品种是否一致
                i = 0
                exit_label = True
                for item_name in self.contract_tables_content["涉案物品返还清单-品种"]:
                    item_quantity = self.contract_tables_content["涉案物品返还清单-规格"][i]
                    i += 1
                    item_for_search = item_name + item_quantity
                    if item_for_search not in other_tabels_content["涉案烟草专卖品核价表-品种规格"]:
                        exit_label = False
                        table_father.display(self,
                                             "品种规格：与《涉案烟草专卖品核价表》不一致，未在此表中查询到的关键词为：" + str(
                                                 item_for_search), "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, '品种',
                                           '品种规格与《涉案烟草专卖品核价表》不一致，未在此表中查询到的关键词为：' + str(
                                               item_for_search))
                        break
                if exit_label:
                    table_father.display(self, "品种种类或数目：正确。与《涉案烟草专卖品核价表》一致", "green")

    def check_date(self):
        """
        作用：返还时间与接收时间一致，且时间与《先行登记保存证据处理通知书》日期一致或之后。
        """
        # 获取返还时间与接收时间
        restitution_time = tyh.get_strtime(self.contract_tables_content["涉案物品返还清单-返还时间"][0])
        reception_time = tyh.get_strtime(self.contract_tables_content["涉案物品返还清单-接收时间"])

        if restitution_time == False or reception_time == False:
            table_father.display(self, "返还时间或接收时间：未填写", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '接收人', '返还时间或接收时间可能未填写，且日期应该具体到XX年XX月XX日')
        elif restitution_time != reception_time:
            table_father.display(self, "返还时间：与接受时间不一致", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '接收人',
                               '返还时间与接受时间不一致，且日期应该具体到XX年XX月XX日,返还时间为：' + str(
                                   restitution_time) + '，接受时间为：' + str(reception_time))
        else:
            # 判断时间是否于《先行登记保存证据处理通知书》日期一致或之后。
            if os.path.exists(self.source_prifix + "先行登记保存证据处理通知书_.docx") == 0:
                table_father.display(self, "文件缺失：《先行登记保存证据处理通知书》不存在", "red")
            else:
                other_info = DocxData(self.source_prifix + "先行登记保存证据处理通知书_.docx")
                other_text = other_info.text
                pattern = re.compile(r'.*(\d{4}年\d{1,2}月\d{1,2}日).*')
                other_time = tyh.get_strtime(re.findall(pattern, other_text)[0])
                if not other_time:
                    table_father.display(self,
                                         "《先行登记保存证据处理通知书》日期：未填写或日期格式有误，日期应该具体到XX年XX月XX日",
                                         "red")
                else:
                    time_differ = tyh.time_differ(restitution_time, other_time)
                    if time_differ >= 0:
                        table_father.display(self,
                                             "返还时间：正确。与接收时间一致，且时间与《先行登记保存证据处理通知书》日期一致或之后",
                                             "green")
                    else:
                        table_father.display(self,
                                             "返还时间：与接收时间一致，但时间与《先行登记保存证据处理通知书》日期不一致或之前",
                                             "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, '返还人',
                                           '【返还时间、接受时间】与《先行登记保存证据处理通知书》日期不一致或之前，且日期应该具体到XX年XX月XX日，《先行登记保存证据处理通知书》日期为：' + str(
                                               other_time))

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
        print("《涉案物品返还清单》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    if "涉案物品返还清单_.docx" in list:
        ioc = Table24(my_prefix, my_prefix)
        contract_file_path = my_prefix + "涉案物品返还清单_.docx"
        ioc.check(contract_file_path, '涉案物品返还清单_.docx')
