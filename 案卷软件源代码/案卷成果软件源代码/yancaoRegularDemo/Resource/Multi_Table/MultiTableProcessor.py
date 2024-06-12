import os
import re
import time

from yancaoRegularDemo.Resource.Multi_Table.people_time_reasonable import table10_people_time_reasonable
from yancaoRegularDemo.Resource.ReadFile import DocxData
from yancaoRegularDemo.Resource.tangyuhao import Precessor1
from yancaoRegularDemo.Resource.table_1_18.table2 import table2
import yancaoRegularDemo.Resource.tools.tanweijia_function as twj
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import win32com

wait_time = 1

maunal_review_item_info = ['不予受理听证通知书_.docx', '立案（不予立案）_.docx', '报告表,延长立案审批表,检查（勘验）笔录_.docx',
                           '证据复制（提取）单_.docx', '公告_.docx', '移送财物清单_.docx', '协助调查函_.docx', '案件调查终结报告_.docx',
                           '涉案物品返还清单_.docx', '听证公告_.docx']


class MultiTableProcessor(table_father):
    def __init__(self, input_dir_dictionary, deviation_percentage=20, output_dir=None, need_to_export=None,
                 progressBar=None,textBrowser = None):
        table_father.__init__(self)
        self.input_dir_dictionary = input_dir_dictionary  # 输入文件路径字典，包含输入的文件夹路径、及各个文件夹下所有文件的路径
        self.output_dir = output_dir.replace("\\", "/")  # 文件夹的输出文件路径
        self.deviation_percentage = deviation_percentage  # 偏差比例值
        self.need_to_export = need_to_export  # 是否需要导出文件到新的文件夹，若为Flase，直接在原文件夹基础上进行修改
        self.progressBar = progressBar
        self.all_info = []  # [[文件夹名，总处理数，错误数目]，[文件夹名，总处理数，错误数目]...]
        self.api_info = {}  # api用的参数

        """
        罚款比例区间，“案件处理审批表”的“承办人意见”中一般都有法律规定的比例，
        比如“40%以上50%以下”  “处以违法货值30%-35%的罚款”
        形式为[7,9],意思是7%以上9%以下的罚款
        该list可以嵌套，如[ [[7,9],[9,11]] , [[30,40],[40,50]] ],分别对应按照案由分组后的各个案件的罚款比例
        """
        self.Penalty_Ratio_Range = []

        """
        案件细节，分别记录 涉案金额、实际罚款比例、罚款金额，类型为字典
        形式为{ 案由1:{文件夹path1:[5000, 7, 350], 文件夹path2:[6000, 8, 480]},案由2:{文件夹path3:[4000, 8, 320], 文件夹path4:[5000, 8, 400]} }
        """
        self.Case_Details = {}

        """
        表10人员——时间合法性
        """
        ioc10 = table10_people_time_reasonable(self.input_dir_dictionary)
        ioc10.check()

        '''
        功能：根据输入的input_dir_dictionary字典，获得多个输入文件夹A、B、C。。的路径input_dir_list
        '''
        input_dir_list = []  # 具体文件的输入文件路径
        for key in self.input_dir_dictionary.keys():
            input_dir_list.append(key)
        self.input_dir_list = input_dir_list  # 输入文件夹A、B、C。。的路径

        '''
        功能：根据输入的多个文件夹A、B、C。。的路径，获得输出文件夹A、B、C。。的路径
        '''
        output_dir_list = []
        # 根据input_dir_dictionary与output_dir，生成输出文件夹路径列表output_dir_list
        for input_dir in self.input_dir_list:
            input_dir.replace("\\", "/")
            output_dir_list.append(self.output_dir + '/' + input_dir[input_dir.rfind('/', 1) + 1:])
        self.output_dir_list = output_dir_list  # 输出文件夹A、B、C。。的路径

        '''
        功能：根据上述计算得到的输出文件夹A、B、C。。的路径，在指定位置创建目录
        改进点：如果文件已存在会报错：FileExistsError: [WinError 183] 当文件已存在时，无法创建该文件。
        '''
        for item in self.output_dir_list:
            try:
                os.mkdir(item)
            except FileExistsError as f:
                print(f)

        '''
        功能：遍历所选中的各个文件夹A、B、C。。中的所有文件，使用Precessor1对这些word文件进行标注,设置进度条，获取错误和总处理的问题数目
        注意：若need_to_export == True，则导出文件到output_dir_list并进行标注，否则直接在原文件夹中进行标注
        '''
        i = 0
        file_count_item = 0  # 用于统计已经加批注了几个文件
        file_count = self.get_file_count(input_dir_dictionary)
        temp2 = {}
        for input_dir in self.input_dir_list:
            temp1 = {}
            for file_dir in self.input_dir_dictionary.get(input_dir):
                # print(file_dir) C:/Users/Xie/Desktop/demo/副本_data2/调查终结报告_.docx
                # 获取处理文件后返回的错误批注信息
                time.sleep(wait_time)
                processor1 = Precessor1(file_dir, output_dir_list[i], self.need_to_export)
                contract_check_result = processor1.action()
                temp1[file_dir] = contract_check_result
                # for item in contract_check_result:
                #     textBrowser.append(item)
                # 获取某文件的total_count和wrong_count,并累加，直到获取该文件夹的总错误和处理数目
                total_count, wrong_count = processor1.processor1_get_count()
                table_father.add_total_count(self, total_count)
                table_father.add_wrong_count(self, wrong_count)
                # 将提示信息写入返回的info_list
                table_father.display(self, "-----------------------------------\n正在审查文件：" + file_dir)
                table_father.display(self, contract_check_result)
                # 用于设置进度条百分比到90%
                file_count_item += 1
                if self.progressBar:
                    self.progressBar.setValue(int(file_count_item / file_count * 90))
                # print(int(file_count_item/file_count*90))
                # 判断此表是否存在需要人工审查项目,若存在，
                file_name = file_dir[file_dir.rfind('/', 1) + 1:]
                if file_name in maunal_review_item_info:
                    table_father.add_manual_item_count(self, 1)
            self.api_info[input_dir] = temp1
            i += 1
            total_count = table_father.get_total_count(self)
            wrong_count = table_father.get_wrong_count(self)
            manual_item_count = table_father.get_manual_item_count(self)
            self.all_info.append([input_dir, total_count, wrong_count,manual_item_count])
            table_father.reset_count(self)
        '''
        遍历所有文件夹的《立案报告表》，根据【案由】，进行文件夹路径的分类
        案由和对应路径列表使用case_dictionary进行存储，key为案由，value为存储路径的list
        若某案由只出现一次，则删除该案由及对应的路径（无法实现对比功能）
        '''
        # case_dictionary = {}
        # for index in range(len(self.input_dir_list)):
        #     # 获取案由
        #     contract_file_path = output_dir_list[index] + "/立案报告表_.docx"
        #     ioc = table2(output_dir_list[index], input_dir_list[index])
        #     cause_of_action = ioc.get_cause_of_action(contract_file_path)
        #     # 判断案由是否在表中，根据判断结果进行更新字典操作
        #     # if not cause_of_action:
        #     #     table_father.display(self,self.input_dir_list[index] + ' 没有立案报告表', 'red')
        #     #     continue
        #     if cause_of_action not in case_dictionary.keys():
        #         key = cause_of_action
        #         case_dictionary[key] = [output_dir_list[index]]
        #     else:
        #         temp = case_dictionary[cause_of_action]
        #         temp.append(output_dir_list[index])
        #         # case_dictionary.update(cause_of_action=temp)
        # # 删除字典中只出现一次的案由
        # for key in list(case_dictionary.keys()):
        #     if len(case_dictionary[key]) <= 1:
        #         del case_dictionary[key]
        #         continue
        # self.case_dictionary = case_dictionary

        '''
        功能：后续各个需执行的函数放入此处，并在self.check()中进行调用
        '''
        self.all_to_check = [
            "self.check_Penalty_Range_Differences(self.progressBar)"
        ]

    '''
    功能：检查 罚款幅度
    '''

    def check_Penalty_Range_Differences(self, progressBar):
        pass
        # 读取“案件处理审批表”的“承办人意见”中法律规定的罚款比例
        # 该list可以嵌套，如[[[7, 9], [9, 11]], [[30, 40], [40, 50]]], 分别对应按照案由分组后的各个案件的罚款比例
        # 92
        # if self.progressBar:
        #     progressBar.setValue(int(92))
        # for key in self.case_dictionary:
        #     path_list = self.case_dictionary[key]
        #     temp_list = []
        #     for i in range(len(path_list)):
        #         data = DocxData(path_list[i] + "/案件处理审批表_.docx")
        #         # 文书中目前发现有两种表达罚款比例的形式
        #         temp_1 = re.search("(\d+)%以上(\d+)%以下", data.tabels_content["承办人意见"])
        #         temp_2 = re.search("(\d+)%-(\d+)%", data.tabels_content["承办人意见"])
        #         if temp_1:
        #             temp_list.append([temp_1.group(1), temp_1.group(2)])
        #         elif temp_2 is not None:
        #             temp_list.append([temp_2.group(1), temp_2.group(2)])
        #         else:
        #             temp_list.append(0)
        #     self.Penalty_Ratio_Range.append(temp_list)
        # print(self.Penalty_Ratio_Range)
        #
        # # 读取“案件处理审批表”的“承办人意见”中的 涉案金额、实际罚款比例、罚款金额
        # """
        # 形式为{ 案由1:{文件夹path1:[5000, 7, 350], 文件夹path2:[6000, 8, 480]},案由2:{文件夹path3:[4000, 8, 320], 文件夹path4:[5000, 8, 400]} }
        # 最后的largest_dict形如：
        # {'涉嫌未在当地烟草专卖批发企业进货': ['D:/tobacco-test/output/副本_data1', 'D:/tobacco-test/output/副本_data2'],
        #  '涉嫌无烟草专卖品准运证运输烟草专卖品': ['D:/tobacco-test/output/副本_data3', 'D:/tobacco-test/output/副本_data4']}
        # """
        # ## 94
        # if self.progressBar:
        #     progressBar.setValue(int(94))
        # largest_dict = {}
        # for key in self.case_dictionary:
        #     # print(key)
        #     large_dict = {}
        #     path_list = self.case_dictionary[key]
        #     for i in range(len(path_list)):
        #         # temp_dict = {}
        #         data = DocxData(path_list[i] + "/案件处理审批表_.docx")
        #         temp = re.search("(\d+).(\d+)元(\d+)%的罚款，计罚款(\d+).(\d+)元", data.tabels_content["承办人意见"])
        #         money_involved = temp[1] + "." + temp[2]
        #         ratio = temp[3]
        #         penalty = temp[4] + "." + temp[5]
        #         detail_list = [money_involved, ratio, penalty]
        #         large_dict.update({path_list[i]: detail_list})
        #
        #     largest_dict.update({key: large_dict})
        # # print(largest_dict)
        #
        # """
        # 按照案由的分组进行打批注，以上一步产生的largest_dict为依据:
        # 一个largest_dict例子:
        # {'涉嫌未在当地烟草专卖批发企业进货': {'D:/tobacco-test/output/副本_data1': ['2365.20', '7', '165.56'], 'D:/tobacco-test/output/副本_data2': ['3317.79', '9', '298.60']},
        # '涉嫌无烟草专卖品准运证运输烟草专卖品': {'D:/tobacco-test/output/副本_data3': ['11404.01', '40', '4561.60'], 'D:/tobacco-test/output/副本_data4': ['12000.40', '40', '4800.16']}}
        # """
        # table_father.display(self, "正在审查同案由案件比较信息：\n")
        # k = 0
        # ## 96
        # if self.progressBar:
        #     progressBar.setValue(int(96))
        # for key in self.case_dictionary:
        #     path_list = self.case_dictionary[key]
        #     for i in range(len(path_list)):
        #         folderPath_to_Remark = path_list[i]
        #         # data = DocxData(folderPath_to_Remark + "/案件处理审批表_.docx")
        #         # loc = data.tabels_content['案由']
        #         for j in range(len(path_list)):
        #             if i == j:
        #                 continue
        #             else:
        #                 content_list = twj.getRemarkContent(path_list[j], largest_dict, key, self.Penalty_Ratio_Range,
        #                                                     k, j)
        #                 # content = "同案由案件" + "(" + content_list[0] + "，" + content_list[1] + ")" + \
        #                 #           "的法定罚款比例区间" + content_list[5] + \
        #                 #           "，其涉案金额、实际罚款比例和罚款金额分别为:" + \
        #                 #           content_list[2] + "元、" + content_list[3] + "%、" + \
        #                 #           content_list[4] + "元"
        #                 content = "同案由案件" + "(" + content_list[0] + "，" + content_list[1] + ")" + \
        #                           "，其涉案金额、实际罚款比例和罚款金额分别为:" + \
        #                           content_list[2] + "元、" + content_list[3] + "%、" + \
        #                           content_list[4] + "元"
        #                 loc_proportion = largest_dict[key][folderPath_to_Remark][1]
        #                 loc = f"元{loc_proportion}%的罚款"
        #                 if not float(loc_proportion) == float(content_list[3]):
        #                     content = content + "，罚款比例与当前案件不一致，请审查。"
        #                 # if self.Penalty_Ratio_Range[k][j] == 0:
        #                 #     pass
        #                 # elif content_list[3] < self.Penalty_Ratio_Range[k][j][0] or content_list[3] > \
        #                 #         self.Penalty_Ratio_Range[k][j][1]:
        #                 #     content = content + "，该同案由案件的罚款比例不在法规规定范围！"
        #                 table_father.display(self, "正在审查" + folderPath_to_Remark + "案件处理审批表_.docx\n")
        #                 table_father.display(self, content)
        #                 twj.addRemarkAboutPenalty(folderPath_to_Remark, '/案件处理审批表_', loc, content)
        #     k = k + 1
        # ## 99
        # if self.progressBar:
        #     progressBar.setValue(int(99))

    def get_file_count(self, input_dir_dictionary):
        num = 0
        for value in input_dir_dictionary.values():
            num = num + len(value)
        return num

    def get_all_info(self):
        return self.all_info

    def check(self):
        for func in self.all_to_check:
            try:
                eval(func)
            except Exception as e:
                table_father.display(self, "文档存在格式错误，函数失效：" + func + ' 遇到错误:' + str(e.args))

        return table_father.get_info_list(self)


if __name__ == '__main__':
    # input_dir_dictionary = {'D:/tobacco-test/input重制3/副本_data1': ['D:/tobacco-test/input重制3/副本_data1/立案报告表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data1/行政处罚决定书_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data1/案件处理审批表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data1/结案报告表_.docx'],
    #                         'D:/tobacco-test/input重制3/副本_data2': ['D:/tobacco-test/input重制3/副本_data2/立案报告表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data2/行政处罚决定书_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data2/案件处理审批表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data2/结案报告表_.docx'],
    #                         'D:/tobacco-test/input重制3/副本_data3': ['D:/tobacco-test/input重制3/副本_data3/立案报告表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data3/行政处罚决定书_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data3/案件处理审批表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data3/结案报告表_.docx'],
    #                         'D:/tobacco-test/input重制3/副本_data4': ['D:/tobacco-test/input重制3/副本_data4/立案报告表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data4/行政处罚决定书_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data4/案件处理审批表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data4/结案报告表_.docx'],
    #                         'D:/tobacco-test/input重制3/副本_data5': ['D:/tobacco-test/input重制3/副本_data5/立案报告表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data5/行政处罚决定书_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data5/案件处理审批表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data5/结案报告表_.docx'],
    #                         'D:/tobacco-test/input重制3/副本_data6': ['D:/tobacco-test/input重制3/副本_data6/立案报告表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data6/行政处罚决定书_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data6/案件处理审批表_.docx',
    #                                                               'D:/tobacco-test/input重制3/副本_data6/结案报告表_.docx']
    #                         }
    input_dir_dictionary = {
        'D:/Demo/副本_data1': ['D:/Demo/副本_data1/立案报告表_.docx', 'D:/Demo/副本_data1/案件处理审批表_.docx'],
        'D:/Demo/副本_data2': ['D:/Demo/副本_data2/立案报告表_.docx', 'D:/Demo/副本_data2/案件处理审批表_.docx']
        # 'D:/Demo/副本_data3': ['D:/Demo/副本_data3/立案报告表_.docx', 'D:/Demo/副本_data3/案件处理审批表_.docx'],
        # 'D:/Demo/副本_data4': ['D:/Demo/副本_data4/立案报告表_.docx', 'D:/Demo/副本_data4/案件处理审批表_.docx'],
        # 'D:/Demo/副本_data5': ['D:/Demo/副本_data5/立案报告表_.docx', 'D:/Demo/副本_data5/案件处理审批表_.docx'],
        # 'D:/Demo/副本_data6': ['D:/Demo/副本_data6/立案报告表_.docx', 'D:/Demo/副本_data6/案件处理审批表_.docx']
    }
    # #
    output_dir = r'D:/output2'
    need_to_export = True
    deviation_percentage = 20
    multiTableProcessor = MultiTableProcessor(input_dir_dictionary, deviation_percentage, output_dir, need_to_export)
    multiTableProcessor.check()
