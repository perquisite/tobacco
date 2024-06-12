# coding:utf-8

import win32com.client

from yancaoRegularDemo.Resource.ReadFile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *
from warnings import simplefilter

simplefilter(action='ignore', category=FutureWarning)

input_dir_dictionary = {'C:/Users/12259/Desktop/原/副本_data6': ['C:/Users/12259/Desktop/原/副本_data6/举报记录表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/先行登记保存批准书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/先行登记保存证据处理通知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/公告_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/协助调查函_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/卷宗封面_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/卷烟鉴别检验样品留样、损耗费用审批表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/延长立案期限告知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/延长立案期限审批表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/延长调查期限告知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/延长调查终结审批表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/抽样取证物品清单_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/指定管辖通知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/撤销立案报告表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/案件处理审批表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/案件移送函_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/案件移送回执_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/案件调查终结报告_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/检查（勘验）笔录_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/涉案烟草专卖品核价表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/涉案物品返还清单_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/移送财物清单_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/立案报告表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/结案报告表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/行政处罚决定书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/证据先行登记保存通知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/证据复制提取单_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/询问笔录_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/询问（调查）通知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data6/调查终结报告_.docx'],
                        'C:/Users/12259/Desktop/原/副本_data2': ['C:/Users/12259/Desktop/原/副本_data2/举报记录表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/先行登记保存批准书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/先行登记保存证据处理通知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/公告_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/协助调查函_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/卷宗封面_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/卷烟鉴别检验样品留样、损耗费用审批表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/延长立案期限告知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/延长立案期限审批表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/延长调查期限告知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/延长调查终结审批表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/抽样取证物品清单_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/指定管辖通知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/撤销立案报告表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/案件处理审批表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/案件移送函_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/案件移送回执_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/案件调查终结报告_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/检查（勘验）笔录_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/涉案烟草专卖品核价表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/涉案物品返还清单_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/移送财物清单_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/立案报告表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/结案报告表_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/行政处罚决定书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/证据先行登记保存通知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/证据复制提取单_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/询问笔录_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/询问（调查）通知书_.docx',
                                                              'C:/Users/12259/Desktop/原/副本_data2/调查终结报告_.docx']}


class table10_people_time_reasonable(table_father):
    def __init__(self, input_dir_dictionary):
        table_father.__init__(self)
        self.input_dir_dictionary = input_dir_dictionary

        self.contract_text = None
        self.contract_tables_content = None
        self.people_time_dict = {}

        self.real_file = []
        self.mw = win32com.client.Dispatch("Word.Application")
        self.doc=None

        self.all_to_check = [
            'self.table10_people_time_reasonable()'

        ]

    def table10_people_time_reasonable(self):
        for key in self.input_dir_dictionary.keys():
            flag = 0
            for file in self.input_dir_dictionary[key]:
                if '询问笔录' in file:
                    self.real_file.append(file)
                    flag = 1
                    break
            if flag == 0:
                print(key + "中不存在讯问笔录")

        print(self.real_file)
        for file in self.real_file:
            print("正在审查" + file + "，审查结果如下：")
            self.mw = win32com.client.Dispatch("Word.Application")
            self.doc = self.mw.Documents.Open(file)

            data = DocxData(file_path=file)
            self.contract_text = data.text
            self.contract_tables_content = data.tabels_content

            self.contract_text = self.strB2Q(self.contract_text)

            if self.gettime(self.contract_text) == False or self.getpeople(self.contract_text) == False:
                table_father.display(self, "时间或人名无法提取", 'red')
                continue
            else:
                time = self.gettime(self.contract_text)
                list0 = self.getpeople(self.contract_text)  # questioner0, questioner1, writer
                # print(self.people_time_dict.keys())
                for l in list0[0:2]:
                    if l not in self.people_time_dict.keys():
                        self.people_time_dict[l] = [time]
                    else:
                        if time in self.people_time_dict[l]:
                            table_father.display(self, "询问人员出现错误，在同一时间检测到有相同的人员在不同的案件中", "red")
                            tyh.addRemarkInDoc(self.mw, self.doc, l, '询问人员出现错误，在同一时间检测到有相同的人员在不同的案件中')
                        else:
                            self.people_time_dict[l].append(time)
        print(self.people_time_dict)

    def strB2Q(self, s):
        """:转："""
        rstring = ""
        for uchar in s:
            if uchar == ":":
                uchar = "："
            rstring += uchar
        return rstring

    def gettime(self, contract_text):
        pattern = ".*询问时间：(.*)"
        time = re.findall(pattern, contract_text)
        if time == [] or time[0].replace(" ", "") == "":
            table_father.display(self, "未查询到 询问时间", "red")
            return False

        else:
            time = time[0]
            if "至" not in time:
                table_father.display(self, "询问时间 格式错误", "red")
                return False

            else:
                time = time.replace(" ", "")
                time0 = time.split("至")[0]
                time1 = time.split("至")[1]
                time0 = tyh.get_strtime_5(time0)
                time1 = tyh.get_strtime_5(time1)
                if time0 == False or time1 == False:
                    table_father.display(self, "询问时间 格式错误", "red")
                    return False
                else:
                    return time0 + "-" + time1

    def getpeople(self, contract_text):

        pattern = ".*询问人：(.*)记录人.*"
        questioner = re.findall(pattern, contract_text)
        # print(questioner)
        if questioner == [] or questioner[0].replace(" ", "") == "":
            table_father.display(self, "未查询到 询问人", "red")
            return False
        else:
            questioner = questioner[0].replace(" ", "")
            # print(questioner)
            questioner0 = questioner.split("、")[0]
            questioner1 = questioner.split("、")[1]
            # print(questioner0)
            # print(questioner1)

        pattern = ".*记录人：(.*).*"
        writer = re.findall(pattern, contract_text)
        if writer == [] or writer[0].replace(" ", "") == "":
            table_father.display(self, "未查询到 记录人", "red")
            return False
        else:
            writer = writer[0].replace(" ", "")
            # print(writer)

        return questioner0, questioner1, writer

    def check(self):
        print("正在审查询问笔录中时间合规性审查与询问人员合法性，审查结果如下：")

        self.table10_people_time_reasonable()
        if self.doc:
            self.doc.Close()
        # self.mw.Quit()
        print("询问笔录审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    ioc = table10_people_time_reasonable(input_dir_dictionary)
    ioc.check()
