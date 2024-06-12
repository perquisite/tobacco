from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh

function_description_dict = {
    'check_compere': '听证主持人不能与《案件调查终结报告》中的调查人一致',
    'check_date': '作出日期与举行听证日期相差7天，即举行听证7日前告知听证人听证时间、地点',
}


# 听证通知书
class Table28(table_father):
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
            "self.check_compere()",
            "self.check_date()"
        ]

    def check_compere(self):
        """
        作用：听证主持人不能与《案件调查终结报告》中的调查人一致
        """
        # if not tyh.file_exists(self.source_prifix, "案件调查终结报告"):
        #     table_father.display(self, "文件缺失：《案件调查终结报告》不存在", "red")
        # else:
        #     other_info = tyh.file_exists_open(self.source_prifix, "案件调查终结报告", DocxData)
        #     other_tabels_content = other_info.tabels_content

    def check_date(self):
        """
        作用：作出日期与举行听证日期相差7天，即举行听证7日前告知听证人听证时间、地点
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
        print("《听证通知书》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result
