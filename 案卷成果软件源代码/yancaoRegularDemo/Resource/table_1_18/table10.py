import os

import win32com.client
import pandas as pd
from yancaoRegularDemo.Resource.ReadFile import *
import re
import shutil
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *
from warnings import simplefilter

simplefilter(action='ignore', category=FutureWarning)

function_description_dict = {
    'timeRight': '“询问时间“不应为空且格式正确',
    'peopleRight': '“询问人“应在记录人名录库中',
}

# 询问笔录
class table10(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix

        self.all_to_check = [
            "self.timeRight()",
            "self.peopleRight()",

        ]

    def timeRight(self):
        pattern = ".*询问时间：(.*)"
        time = re.findall(pattern, self.contract_text)
        if time == [] or time[0].replace(" ", "") == "":
            table_father.display(self, "询问时间：未查询到 询问时间", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '询问时间：', '未查询到 询问时间')
        else:
            time = time[0]
            if "至" not in time:
                table_father.display(self, "询问时间：询问时间 格式错误", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '询问时间：', '询问时间 格式错误')
            else:
                time = time.replace(" ", "")
                time0 = time.split("至")[0]
                time1 = time.split("至")[1]
                time0 = tyh.get_strtime_5(time0)
                time1 = tyh.get_strtime_5(time1)
                if time0 == False or time1 == False:
                    table_father.display(self, "询问时间：询问时间 格式错误", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '询问时间：', '询问时间 格式错误')

    def peopleRight(self):
        # print(tyh.read_txt(self.source_prifix+"名录库-询问人.txt"))
        # questioner_real = tyh.read_txt(self.source_prifix + "名录库-询问人.txt")
        if not os.path.exists(self.source_prifix+'data'):
            src_data_dir = str(os.getcwd()).replace('\\','/')+'/Resource/data'
            # print(src_data_dir)
            shutil.copytree(src_data_dir, self.source_prifix+'data')
            table_father.display(self, "名录库缺失：案卷中不存在【data】文件夹，已在输入文件夹中自动生成相关文件夹及文档。默认生成的文档为眉山市【现场执法人员名录.xls】", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '询问笔录', '名录库缺失：案卷中不存在【data】文件夹，已在输入文件夹中自动生成相关文件夹及文档。默认生成的文档为眉山市【现场执法人员名录.xls】')
        if not os.path.exists(self.source_prifix+'data/现场执法人员名录.xls'):
            shutil.copy(src_data_dir+'/现场执法人员名录.xls',self.source_prifix+'data')
            table_father.display(self,
                                 "名录库缺失：【data】文件夹缺失“现场执法人员名录.xls”，已在输入文件夹中自动生成默认文件",
                                 "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '询问笔录',
                               '名录库缺失：【data】文件夹缺失“现场执法人员名录.xls，已在输入文件夹中自动生成默认文件”')
        if not os.path.exists(self.source_prifix+'data/名录库-记录人.txt'):
            shutil.copy(src_data_dir+'/名录库-记录人.txt', self.source_prifix + 'data')
            table_father.display(self,
                                 "名录库缺失：【data】文件夹缺失“名录库-记录人.txt”，已在输入文件夹中自动生成默认文件",
                                 "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '询问笔录',
                               '名录库缺失：【data】文件夹缺失“名录库-记录人.txt”，已在输入文件夹中自动生成默认文件')
        if not os.path.exists(self.source_prifix+'data/名录库-询问人.txt'):
            shutil.copy(src_data_dir+'/名录库-询问人.txt', self.source_prifix + 'data')
            table_father.display(self,
                                 "名录库缺失：【data】文件夹缺失“名录库-询问人.txt”，已在输入文件夹中自动生成默认文件",
                                 "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '询问笔录',
                               '名录库缺失：【data】文件夹缺失“名录库-询问人.txt”，已在输入文件夹中自动生成默认文件')

        read_questioner_real = pd.read_excel(io=self.source_prifix+'data\\现场执法人员名录.xls', usecols=[2], names=None)
        questioner_real = []
        for name_ in read_questioner_real.values.tolist():
            questioner_real.append(name_[0])
        writer_real = tyh.read_txt(self.source_prifix + "data\\名录库-记录人.txt")

        pattern = ".*询问人：(.*)记录人.*"
        questioner = re.findall(pattern, self.contract_text)
        # print(questioner)
        if questioner == [] or questioner[0].replace(" ", "") == "":
            table_father.display(self, "询问人：未查询到 询问人", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '询问人：', '未查询到 询问人')
        else:
            questioner = questioner[0].replace(" ", "")
            # print(questioner)
            questioner0 = questioner.split("、")[0]
            questioner1 = questioner.split("、")[1]
            # print(questioner0)
            # print(questioner1)
            if questioner0 not in questioner_real:
                table_father.display(self, "询问人名录库"+str(questioner0) + "未在 询问人名录库中", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, questioner0, str(questioner0) + "未在 询问人名录库中")
            else:
                table_father.display(self, "询问人名录库"+str(questioner0) + "在 询问人名录库中", "green")

            if questioner1 not in questioner_real:
                table_father.display(self, "询问人名录库"+str(questioner1) + "未在 询问人名录库中", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, questioner1, str(questioner1) + "未在 询问人名录库中")
            else:
                table_father.display(self, "询问人名录库"+str(questioner1) + "在 询问人名录库中", "green")

        pattern = ".*记录人：(.*).*"
        writer = re.findall(pattern, self.contract_text)
        if writer == [] or writer[0].replace(" ", "") == "":
            table_father.display(self, "记录人：未查询到 记录人", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '记录人：', '未查询到 记录人')
        else:
            writer = writer[0].replace(" ", "")
            # print(writer)
            if writer not in writer_real:
                table_father.display(self, "询问人名录库"+str(writer) + "未在 记录人名录库中", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, writer, str(writer) + "未在 记录人名录库中")
            else:
                table_father.display(self, "询问人名录库"+str(writer) + "在 记录人名录库中", "green")

    def check(self, contract_file_path,file_name_real):
        print("正在审查"+file_name_real+"，审查结果如下：")
        self.mw = win32com.client.Dispatch("Word.Application")
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
        print(file_name_real+"审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\in\\"
    list = os.listdir(my_prefix)
    if "询问笔录_.docx" in list:
        ioc = table10(my_prefix, my_prefix)
        contract_file_path = os.path.join(my_prefix, "询问笔录_.docx")
        ioc.check(contract_file_path, "询问笔录_.docx")
