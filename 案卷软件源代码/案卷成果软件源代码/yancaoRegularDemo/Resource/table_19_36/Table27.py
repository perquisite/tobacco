import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
from yancaoRegularDemo.Resource.tools.utils import chinese_to_date
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
import os

function_description_dict = {
    'check_date': '作出日期与《行政处罚事先告知书》作出日期保持一致或在之前。',
    'check_consistency': '案发时间、违法行为、违法条款、处罚依据、拟行政处罚与《案件调查终结报告》保持一致。',
}


# 听证告知书
class Table27(table_father):
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
            "self.check_date()",
            "self.check_consistency()"
        ]

    def check_date(self):
        """
        作用：作出日期与《行政处罚事先告知书》作出日期保持一致或在之前。
        """
        if not tyh.file_exists(self.source_prifix, "行政处罚事先告知书"):
            table_father.display(self, "文件缺失：《行政处罚事先告知书》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "行政处罚事先告知书", DocxData)
            other_tabels_text = other_info.text
            time_pattern = re.compile(r'.*(二.{3}年.{1,2}月.{1,3}日).*')
            this_file_time = re.findall(time_pattern, self.contract_text)[0]
            other_file_time = re.findall(time_pattern, other_tabels_text)[0]
            if this_file_time == '' or this_file_time is None:
                table_father.display(self, "作出日期：未找到作出日期，作出日期请采用中文格式的年月日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '听证告知书', '未找到作出日期，作出日期请采用中文格式的年月日')
            elif other_file_time == '' or other_file_time is None:
                table_father.display(self, "作出日期：《行政处罚事先告知书》未找到作出日期，作出日期请采用中文格式的年月日",
                                     "red")
            else:
                this_file_time_date = chinese_to_date(this_file_time)
                other_file_time_date = chinese_to_date(other_file_time)
                time_differ = tyh.time_differ(this_file_time_date, other_file_time_date)
                if time_differ <= 0:
                    table_father.display(self, "作出日期：正确。与《行政处罚事先告知书》作出日期保持一致或在之前", "green")
                else:
                    table_father.display(self, "作出日期：未与《行政处罚事先告知书》作出日期保持一致或之前", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, this_file_time,
                                       '作出日期错误，未与《行政处罚事先告知书》作出日期保持一致或之前,《行政处罚事先告知书》作出日期为：' + str(
                                           other_file_time_date))

    def check_consistency(self):
        """
        作用：案发时间、违法行为、违法条款、处罚依据、拟行政处罚与《案件调查终结报告》保持一致。
            a“案发时间”识别：“你（单位）于X年X月X日”。与《案件调查终结报告》中“X年X月X日，我局执法人员XXX”保持一致。
            b“违法行为”识别：“因XXX的行为”。与《案件调查终结报告》案由一栏保持一致。
            c“违法条款”识别：“违反了XXX的规定。” 与《案件调查终结报告》中“违反	XXX”保持一致。
            d“处罚依据”识别：“依据XXX”。与《案件调查终结报告》中“依据XXX”保持一致。
            e“拟行政处罚”识别：“本局拟对你（单位）作出XXX的行政处罚”。与《案件调查终结报告》中“建议作如下行政处罚：XXXXX”保持一致。
        """
        if not tyh.file_exists(self.source_prifix, "案件调查终结报告"):
            table_father.display(self, "文件缺失：《案件调查终结报告》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "案件调查终结报告", DocxData)
            other_tabels_content = other_info.tabels_content

            # a“案发时间”识别：“你（单位）于X年X月X日”。与《案件调查终结报告》中“X年X月X日，我局执法人员XXX”保持一致。
            time_parttern1 = re.compile(r'你（单位）于(\d{4}年\d{1,2}月\d{1,2}日)因')
            time_parttern2 = re.compile(r'(\d{4}年\d{1,2}月\d{1,2}日).*我局执法人员')

            time1 = re.findall(time_parttern1, self.contract_text)[0]
            time2 = re.findall(time_parttern2, other_tabels_content['调查事实'])[0]

            if time1 != time2:
                table_father.display(self, "案发时间：与《案件调查终结报告》不一致,请检查，并参考文档检查格式", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "你（单位）于" + str(time1) + "因",
                                   '案发时间与《案件调查终结报告》不一致,请检查，并参考文档检查格式。《案件调查终结报告》时间为:' + str(
                                       time2))
            else:
                table_father.display(self, "案发时间：正确。与《案件调查终结报告》一致", "green")

            # b“违法行为”识别：“因XXX的行为”。与《案件调查终结报告》案由一栏保持一致。
            behavior_parttern = re.compile(r'因(.*)的行为')
            behavior = re.findall(behavior_parttern, self.contract_text)[0]
            other_behavior = other_tabels_content['案由']
            if behavior != other_behavior:
                table_father.display(self, "案由：与《案件调查终结报告》不一致,请检查，并参考文档检查格式", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "因" + str(behavior) + "的行为",
                                   '案由与《案件调查终结报告》不一致,请检查，并参考文档检查格式。《案件调查终结报告》案由为：' + str(
                                       other_behavior))
            else:
                table_father.display(self, "案由：正确。与《案件调查终结报告》一致", "green")

            # c“违法条款”识别：“违反了XXX的规定。” 与《案件调查终结报告》中“违反 XXX”保持一致。
            clause_parttern = re.compile(r'违反了(.*)的规定')
            other_clause_parttern = re.compile(r'违反了(.*)[之的]规定')
            clause = re.findall(clause_parttern, self.contract_text)[0]
            other_clause = re.findall(other_clause_parttern, other_tabels_content['案件性质'])[0]
            if clause != other_clause:
                table_father.display(self, "违法条款：与《案件调查终结报告》不一致,请检查，并参考文档检查格式", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "违反了" + str(clause) + "的规定",
                                   '违法条款与《案件调查终结报告》不一致,请检查，并参考文档检查格式。《案件调查终结报告》描述为：' + str(
                                       other_clause))
            else:
                table_father.display(self, "违法条款：正确。与《案件调查终结报告》一致", "green")

            # d“处罚依据”识别：“依据XXX”。与《案件调查终结报告》中“依据XXX”保持一致。
            gist_parttern = re.compile(r'依据(.*)，本局拟对你')
            gist = re.findall(gist_parttern, self.contract_text)[0]
            other_gist = other_tabels_content['处罚依据']
            if gist != other_gist:
                table_father.display(self, "处罚依据：与《案件调查终结报告》不一致,请检查，并参考文档检查格式", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "依据" + gist + "，本局拟对你",
                                   '处罚依据与《案件调查终结报告》不一致,请检查，并参考文档检查格式。《案件调查终结报告》处罚依据为：' + str(
                                       other_gist))
            else:
                table_father.display(self, "处罚依据：正确。与《案件调查终结报告》一致", "green")

            # e“拟行政处罚”识别：“本局拟对你（单位）作出XXX的行政处罚”。与《案件调查终结报告》中“建议作如下行政处罚：XXXXX”保持一致。
            punishment_parttern = re.compile(r'本局拟对你（单位）作出(.*)的行政处罚')
            other_punishment_parttern = re.compile(r'建议作如下行政处罚：(.*)')
            punishment = re.findall(punishment_parttern, self.contract_text)[0]
            other_punishment = re.findall(other_punishment_parttern, other_tabels_content['处理意见'])
            if other_punishment == []:
                tyh.addRemarkInDoc(self.mw, self.doc, "听证告知书",
                                   "未在《案件调查终结报告》中匹配到行政处罚建议，匹配模式为：建议作如下行政处罚：xxx")
            elif punishment != other_punishment[0]:
                table_father.display(self, "行政处罚建议：与《案件调查终结报告》不一致,请检查，并参考文档检查格式", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, "本局拟对你（单位）作出" + str(punishment[0]) + "的行政处罚",
                                   '行政处罚建议与《案件调查终结报告》不一致,请检查，并参考文档检查格式。《案件调查终结报告》行政处罚建议为：' + str(
                                       other_punishment[0]))
            else:
                table_father.display(self, "行政处罚建议：正确。与《案件调查终结报告》一致", "green")

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
        print("《听证告知书》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    if "听证告知书_.docx" in list:
        ioc = Table27(my_prefix, my_prefix)
        contract_file_path = my_prefix + "听证告知书_.docx"
        ioc.check(contract_file_path, '听证告知书_.docx')
