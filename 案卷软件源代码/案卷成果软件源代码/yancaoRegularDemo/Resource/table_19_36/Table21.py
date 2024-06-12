import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
import datetime
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import get_nearest_working_day

function_description_dict = {
    'check_difference': '案由、立案日期、当事人、证件类型及号码、地址、联系电话、案情摘要一般应与《立案报告表》保持一致',
    'check_date_legal': '承办人签名日期、承办部门意见日期、负责人意见日期应在《立案报告表》负责人意见日期一栏30天内。('
                        '小于等于二十九天)。若限期届满之日为法定节假日（比如周末、五一假），则顺延至节假日之后的第一天。',
}

# 延长调查终结审批表
class Table21(table_father):
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
            "self.check_difference()",
            "self.check_date_legal()"
        ]

    def check_difference(self):
        """
        作用：案由、立案日期、当事人、证件类型及号码、地址、联系电话、案情摘要一般应与《立案报告表》保持一致，
        若出现经调查变动的栏目，比如变更了当事人、案由，出现预警提示。
        """
        if not tyh.file_exists(self.source_prifix, "立案报告表"):
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            register_info = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
            register_tabels_content = register_info.tabels_content
            ready_to_check = ["案由", "立案日期", "当事人", "证件类型及号码", "地址", "联系电话", "案情摘要"]
            for i in ready_to_check:
                if i == "立案日期":
                    # 获取立案报告表的立案日期并处理
                    date_pattern = re.compile(r'.*日期：(.*)')
                    date_by_principal = re.findall(date_pattern, register_tabels_content["负责人意见"].replace(":", "："))
                    if date_by_principal == [''] or date_by_principal == []:
                        table_father.display(self, "立案报告表_负责人意见_日期：应具体到XX年XX月XX日", "red")
                    date_by_principal = tyh.get_strtime(date_by_principal[0])
                    # 获取延长调查终结审批表的立案日期并处理
                    # 将二者进行比较
                    if date_by_principal == tyh.get_strtime(self.contract_tables_content["立案日期"]):
                        table_father.display(self, "立案日期：正确。与《立案报告表》立案日期（即负责人意见签名时间）一致", "green")
                    else:
                        table_father.display(self, "立案日期：与《立案报告表》立案日期（即负责人意见签名时间）不一致", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, '立案日期', '立案日期与《立案报告表》立案日期（即负责人意见签名时间）不一致，《立案报告表》立案日期为：'+date_by_principal)
                elif i == "联系电话":
                    pass
                else:
                    if register_tabels_content[i] == self.contract_tables_content[i]:
                        table_father.display(self, i + "：正确。与《立案报告表》一致,《立案报告表》中为："+register_tabels_content[i], "green")
                    else:
                        table_father.display(self, i + "：与《立案报告表》不一致,《立案报告表》中为："+register_tabels_content[i], "red")
                        if i == "案由":
                            tyh.addRemarkInDoc(self.mw, self.doc, "案   由", i + "与《立案报告表》不一致,《立案报告表》中为："+register_tabels_content[i])
                        elif i == "地址":
                            tyh.addRemarkInDoc(self.mw, self.doc, "地   址", i + "与《立案报告表》不一致,《立案报告表》中为："+register_tabels_content[i])
                        else:
                            tyh.addRemarkInDoc(self.mw, self.doc, i, i+"与《立案报告表》不一致,《立案报告表》中为："+register_tabels_content[i])

    def check_date_legal(self):
        """
        作用：承办人签名日期、承办部门意见日期、负责人意见日期应在《立案报告表》负责人意见日期一栏30天内。(小于等于二十九天)
        若限期届满之日为法定节假日（比如周末、五一假），则顺延至节假日之后的第一天。建议时间超期，出现预警提示
        """
        # 判断限期届满之日是否为法定节假日，是的话将该日期改为节假日之后的第一天
        if not tyh.file_exists(self.source_prifix, "立案报告表"):
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            register_info = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
            register_tabels_content = register_info.tabels_content
            date_pattern = re.compile(r'.*日期：(.*)')
            date_by_principal = re.findall(date_pattern, register_tabels_content["负责人意见"].replace(":", "："))
            if date_by_principal == [''] or date_by_principal == []:
                table_father.display(self, "立案报告表_负责人意见_日期：应具体到XX年XX月XX日", "red")
            date_by_principal = tyh.get_strtime(date_by_principal[0])
            date_by_principal_plus_days = (
                    datetime.datetime.strptime(date_by_principal, '%Y-%m-%d') + datetime.timedelta(days=29)).date()
            # 最近的工作日期
            nearest_working_day = get_nearest_working_day(date_by_principal_plus_days)

            # 获取承办人签名日期
            '''
            表中无承办人签名日期，只有延长调查终结事由及期限的日期，暂用此日期替代
            '''
            # text_undertaker = str(self.contract_tables_content["承办人意见"]).replace(":", "：")
            # date_by_undertaker = re.findall(date_pattern, text_undertaker)
            #
            # if date_by_undertaker == [''] or date_by_undertaker == []:
            #     table_father.display(self, "× 承办人签名日期_日期 应具体到XX年XX月XX日", "red")
            #     tyh.addRemarkInDoc(self.mw, self.doc, '承办部门意见', '日期应具体到XX年XX月XX日')
            # else:
            #     date_by_undertaker = tyh.get_strtime(date_by_undertaker[0])
            #     time_differ = tyh.time_differ(nearest_working_day, date_by_undertaker)
            #     if time_differ >= 0:
            #         table_father.display(self, "√ 承办人意见_日期 在《立案报告表》负责人意见日期一栏30天内(如有法定节假日，顺延到法定节假日后一天)", "green")
            #     else:
            #         table_father.display(self, "× 承办人意见_日期 不在《立案报告表》负责人意见日期一栏30天内(如有法定节假日，顺延到法定节假日后一天)", "red")
            #         tyh.addRemarkInDoc(self.mw, self.doc, '承办部门意见', '承办人日期不在《立案报告表》负责人意见日期一栏30天内(如有法定节假日，顺延到法定节假日后一天)')

            # 获取承办部门意见日期
            text_department = str(self.contract_tables_content["承办部门意见"]).replace(":", "：")
            date_by_department = re.findall(date_pattern, text_department)

            if date_by_department == [''] or date_by_department == []:
                table_father.display(self, "承办部门意见_日期：应具体到XX年XX月XX日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办部门意见', '日期应具体到XX年XX月XX日')
            else:
                date_by_department = tyh.get_strtime(date_by_department[0])
                time_differ = tyh.time_differ(nearest_working_day, date_by_department)
                if time_differ >= 0:
                    table_father.display(self, "承办部门意见_日期：正确。在《立案报告表》负责人意见日期一栏30天内(如有法定节假日，顺延到法定节假日后一天)", "green")
                else:
                    table_father.display(self, "承办部门意见_日期：不在《立案报告表》负责人意见日期一栏30天内(如有法定节假日，顺延到法定节假日后一天)", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '承办部门意见', '日期不在《立案报告表》负责人意见日期一栏30天内(如有法定节假日，顺延到法定节假日后一天),《立案报告表》负责人意见为：'+date_by_principal)

            # 获取负责人意见日期
            text_master = str(self.contract_tables_content["负责人意见"]).replace(":", "：")
            date_by_master = re.findall(date_pattern, text_master)

            if date_by_master == [''] or date_by_master == []:
                table_father.display(self, "负责人意见_日期：应具体到XX年XX月XX日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '日期应具体到XX年XX月XX日')
            else:
                date_by_master = tyh.get_strtime(date_by_master[0])
                time_differ2 = tyh.time_differ(nearest_working_day, date_by_master)
                if time_differ2 >= 0:
                    table_father.display(self, "负责人意见_日期：正确。在《立案报告表》负责人意见日期一栏30天内(如有法定节假日，顺延到法定节假日后一天)", "green")
                else:
                    table_father.display(self, "负责人意见_日期：不在《立案报告表》负责人意见日期一栏30天内(如有法定节假日，顺延到法定节假日后一天)", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '日期不在《立案报告表》负责人意见日期一栏30天内(如有法定节假日，顺延到法定节假日后一天),《立案报告表》负责人意见为：'+date_by_principal)

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
        print("《延长调查终结审批表》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    if "延长案件调查终结审批表_.docx" in list:
        ioc = Table21(my_prefix, my_prefix)
        contract_file_path = my_prefix + "延长案件调查终结审批表_.docx"
        ioc.check(contract_file_path,'延长案件调查终结审批表_.docx')
