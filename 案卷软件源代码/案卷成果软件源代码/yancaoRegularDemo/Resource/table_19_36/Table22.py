import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import chinese_to_date

function_description_dict = {
    'check_name': '被告知人名字/名称一般与《立案报告表》当事人保持一致',
    'check_date': '作出时间与《延长案件调查终结审批表》做出时间一致或之后。',
}

# 延长调查期限告知书
class Table22(table_father):
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
            "self.check_name()",
            "self.check_date()"
        ]

    def check_name(self):
        """
        作用：被告知人名字/名称一般与《立案报告表》当事人保持一致，若经调查出现变更当事人，出现预警提示
        """
        if not tyh.file_exists(self.source_prifix, "立案报告表"):
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            register_info = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
            register_tabels_content = register_info.tabels_content
            # 获取当事人姓名及姓名长度
            register_name = register_tabels_content["当事人"]
            register_name_len = len(register_name)

            # 获取被告知人姓名
            name_pattern_str = "(.{" + str(register_name_len) + "})：\n你（单位）"
            name_pattern = re.compile(name_pattern_str)
            nunciatus_name = re.findall(name_pattern, self.contract_text)

            if nunciatus_name is not [] and nunciatus_name == register_name:
                table_father.display(self, "被告知人名字/名称：正确。与《立案报告表》当事人保持一致", "green")
            else:
                table_father.display(self, "被告知人名字/名称：与《立案报告表》当事人不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, str(nunciatus_name[-2:]) + '：', '被告知人名字/名称与《立案报告表》当事人不一致，《立案报告表》当事人为：'+register_name)

    def check_date(self):
        """
        作用：作出时间与《延长案件调查终结审批表》做出时间一致或之后。
        """
        if not tyh.file_exists(self.source_prifix, "延长案件调查终结审批表"):
            table_father.display(self, "文件缺失：《延长案件调查终结审批表》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "延长案件调查终结审批表", DocxData)
            other_tabels_content = other_info.tabels_content
            time_pattern = re.compile(r'.*(二.{3}年.{1,2}月.{1,3}日).*')
            other_time_pattern = re.compile(r'.*日期：(.*)')
            this_file_time = re.findall(time_pattern, self.contract_text)
            other_file_time = re.findall(other_time_pattern, other_tabels_content['延长调查终结事由及期限'])
            if this_file_time[0] == '' or this_file_time[0] is None or this_file_time[0] == []:
                table_father.display(self, "作出日期：未找到作出日期，作出日期请采用中文格式的年月日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '延长调查期限告知书', '未找到作出日期，作出日期请采用中文格式的年月日')
            elif other_file_time[0] == '' or other_file_time[0] is None or other_file_time[0] == []:
                table_father.display(self, "他表做出日期：《延长调查终结审批表》未找到作出日期，作出日期请采用中文格式的年月日", "red")
            else:
                this_file_time_date = chinese_to_date(this_file_time[0])
                other_file_time_date  = tyh.get_strtime(other_file_time[0])
                time_differ = tyh.time_differ(this_file_time_date, other_file_time_date)
                if time_differ < 0:
                    table_father.display(self, "作出日期：作出日期应在《延长调查终结审批表》做出时间一致或之后", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, this_file_time[0], "作出日期应在《延长调查终结审批表》做出时间一致或之后,《延长调查终结审批表》做出时间为："+other_file_time[0])
                else:
                    table_father.display(self, "作出日期：正确。作出日期与《延长调查终结审批表》做出日期一致或之后", "green")

            """
            作用：签收日期与本文书作出时间一致或之后
            """
            tyh.addRemarkInDoc(self.mw, self.doc, this_file_time[0], "签收日期应该与本文书作出时间一致或之后，请主观审查")

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
            # eval(func)
        self.doc.Save()
        self.doc.Close()
        
        self.mw.Quit()
        # self.mw.Quit()
        print("《延长调查期限告知书》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    if "延长调查期限告知书_.docx" in list:
        ioc = Table22(my_prefix, my_prefix)
        contract_file_path = my_prefix + "延长调查期限告知书_.docx"
        ioc.check(contract_file_path, '延长调查期限告知书_.docx')
