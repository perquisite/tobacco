import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh

function_description_dict = {
    'check_case': '案由、案件编号与《立案报告表》保持一致。',
    'check_date': '签名日期与陈述、申辩时间保持一致',
    'check_sign': '参与陈述、申辩的执法人员应为两名以上执法人员',
}

# 陈述申辩记录
class Table26(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prifix = source_prifix# 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.mw.Visible = 0
        self.mw.DisplayAlerts = 0
        self.contract_text = None
        self.contract_tables_content = None
        

        self.all_to_check = [
            "self.check_case()",
            "self.check_date()",
            "self.check_sign()"
        ]

    

    def check_case(self):
        """
        作用：案由、案件编号与《立案报告表》保持一致。
        """
        # 未找到《立案报告表》
        if not tyh.file_exists(self.source_prifix, "立案报告表"):
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
            other_tabels_content = other_info.tabels_content
            other_cause_of_action = other_tabels_content['案由']

            # 获取案由
            cause_of_action_parttern = re.compile(r'案由：(.*)[\s]*')
            cause_of_action = re.findall(cause_of_action_parttern, self.contract_text)[0]
            if cause_of_action == '' or cause_of_action is None:
                table_father.display(self,"案由：案由为空", "red")

            # 获取案件编号
            case_number_parttern = re.compile(r'案件编号：(.*)[\s]*')
            case_number = re.findall(case_number_parttern, self.contract_text)[0]
            if case_number == '' or case_number is None:
                table_father.display(self,"案件编号：案件编号为空", "red")

            if cause_of_action == other_cause_of_action:
                table_father.display(self,"案由：正确。与《立案报告表》保持一致", "green")
            else:
                table_father.display(self,"案由：与《立案报告表》不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案由', '案由与《立案报告表》不一致,《立案报告表》案由为：'+str(other_cause_of_action))
            if case_number not in other_info.text:
                table_father.display(self, "案件编号：与《立案报告表》不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案件编号', '案件编号与《立案报告表》不一致')


    def check_date(self):
        """
        作用：签名日期与陈述、申辩时间保持一致
        """
        time_parttern = re.compile(r'.*(\d{4}年\d{1,2}月\d{1,2}日)')
        statement_time = tyh.get_strtime(re.findall(time_parttern, self.contract_text)[0])
        sign_time = tyh.get_strtime(re.findall(time_parttern, self.contract_text)[0])

        if statement_time is not False and statement_time == sign_time:
            table_father.display(self,"签名日期：正确。与陈述、申辩时间保持一致", "green")
        else:
            table_father.display(self,"签名日期：与陈述、申辩时间不一致", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, re.findall(time_parttern, self.contract_text)[2], '签名日期与陈述、申辩时间不一致,陈述、申辩时间为：'+str(statement_time))

    def check_sign(self):
        """
        作用：参与陈述、申辩的执法人员应为两名以上执法人员
        """
        table_father.display(self, "陈述、申辩人（签名）：提示。参与陈述、申辩的执法人员应为两名以上执法人员,请主观审查", "red")
        tyh.addRemarkInDoc(self.mw, self.doc, '陈述、申辩人（签名）', '参与陈述、申辩的执法人员应为两名以上执法人员,请主观审查')

    def check(self, contract_file_path,file_name_real):
        print("正在审查"+file_name_real+"，审查结果如下：")
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
        print("《陈述申辩记录》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:/Users/Xie/Desktop/out/"
    list = os.listdir(my_prefix)
    if "陈述申辩记录_.docx" in list:
        ioc = Table26(my_prefix, my_prefix)
        contract_file_path = my_prefix + "陈述申辩记录_.docx"
        ioc.check(contract_file_path, '陈述申辩记录_.docx')