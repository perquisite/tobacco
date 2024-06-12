import win32com
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

function_description_dict = {
    'warning': '主观审查不予受理听证理由。',
}


# 不予受理听证通知书
class Table29(table_father):
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
            "self.warning()"
        ]

    def warning(self):
        """
        作用：主观审查不予受理听证理由。
        """
        table_father.display(self, "不予受理听证理由：提示。请主观审查不予受理听证理由", "red")

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")

        # self.contract_text = data.text
        # self.contract_tables_content = data.tabels_content
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
