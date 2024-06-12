import rdflib
import json
import pandas as pd
import os
import numpy as np

from zhishitupu.src.function import is_list_consecutive
from zhishitupu.tools.simple_content import Simple_Content


def check_excel(file):
    data = pd.read_excel(file)
    all_to_check = [
        "check_null_1(data)",
        "check_null_2(data)",
        "check_null_3(data)",
        "check_penalty(data)"
        # "check_result_num(data)"

    ]
    no_blank = check_null_0(data)
    if not no_blank:
        return "表中有空白未填。"
    else:
        r = []
        for func in all_to_check:
            try:
                temp = eval(func)
                if temp:
                    r += temp
            except Exception as e:
                # print(str(e.args))
                tip = 'Excel表中依然有未知错误。请再次检查填写内容！。'
                # print(tip)
                return tip
        return r


# 1 “最小处罚比例”或“最大处罚比例”不为数字，则报错
# 2 若“罚款”为“是”，该栏对应“最小（最大）惩罚比例”不能都为空（“罚款”只能为“是”、“否”和“无”）
def check_penalty(data):
    info_list = []
    penalty = data['罚款'][1:].tolist()
    i = 1
    while i <= len(penalty):
        if "是" in penalty[i-1]:
            min_p = str(data['最小惩罚'][i]).strip()
            max_p = str(data['最大惩罚'][i]).strip()
            # print(min_p, max_p)
            if str(min_p) in ["nan", "", "无"] and str(max_p) in ["nan", "", "无"]:
                info_list.append('若“罚款”为“是”，则该栏对应“最小（最大）惩罚”不能都为空！')
            else:
                sc = Simple_Content()
                for j in [min_p, max_p]:
                    if not sc.is_number(j):
                        info_list.append('“最小处罚比例”或“最大处罚比例”不为数字！')
                        break
        i += 1

    return info_list


# “处罚结果”对应“处罚结果原文”不为“无”
def check_null_3(data):
    info_list = []
    i = 1
    for item in data['处罚结果原文'][1:data.index.stop]:
        if i==1:
            continue
        temp = str(item).strip()
        if not temp == "nan" and temp == "无":
            info_list.append("各个“处罚结果”对应的“处罚结果原文“不应为“无”")
            return info_list
        i += 1


# “违法行为”对应“法律条款原文”、“违法行为名称”、“法律名称”、“法律条数”不为“无”
def check_null_2(data):
    info_list = []
    index = data[data['条目'] == '违法行为'].index.tolist()[0]
    text = str(data['法律条款原文'][index]).strip()
    behavior_name = str(data['违法行为名称'][index]).strip()
    law_name = str(data['法律名称'][index]).strip()
    law_num = str(data['法律条数'][index]).strip()
    # print(text)
    # print(behavior_name)
    # print(law_name)
    # print(law_num)
    if not text == "nan" and text == "无":
        info_list.append("“法律条款原文”不应为“无”")
    if not behavior_name == "nan" and behavior_name == "无":
        info_list.append("“违法行为名称”不应为“无”")
    if not law_name == "nan" and law_name == "无":
        info_list.append("“法律名称”不应为“无”")
    if not law_num == "nan" and law_num == "无":
        info_list.append("“法律条数”不应为“无”")
    # print(info_list)
    return info_list


# “罚款”、“没收烟草或烟草制品”、“没收违法所得”和“收购”四列数据只能为“是”、“否”和“无”
def check_null_1(data):
    info_list = []
    for i in [data.columns[3], data.columns[6], data.columns[7], data.columns[11]]:
        temp = data[i].values
        # print(temp)
        for j in temp:
            # print(j == "nan")
            if str(j) not in ["是", "否", "无", "nan"]:
                info_list.append("“罚款”、“没收烟草或烟草制品”、“没收违法所得”和“收购”所在列的内容只能为“是”、“否”或“无”")
                return info_list


# 有空白未填(空格和nan两种)
def check_null_0(data):
    # info_list = []
    for i in data.columns[1:]:
        temp = data[i].values
        for j in temp:
            # print(str(j))
            if str(j) == "nan" or str(j).strip() == "":
                # info_list.append("表中有空白未填。")
                return False
    return True


"""
# “处罚结果”编号不连贯或缺失
# “金额阈值”若有n个，“处罚结果”为n+1个
def check_result_num(data):
    info_list = []
    temp = data["金额阈值"][0].strip()
    l1 = temp.split()
    yuzhi_num = len(l1)
    temp = [i.strip() for i in data["条目"][1:].tolist()]
    resultNo_list = [int(i.strip("处罚结果")) for i in temp]
    result_num = len(resultNo_list)
    if not result_num == yuzhi_num+1:
        info_list.append("金额阈值于处罚结果数目不对应。")
    if not is_list_consecutive(resultNo_list):
        info_list.append("处罚结果编号不连贯或缺失。")

    return info_list
"""

if __name__ == "__main__":
    pth = r'C:\Users\twj\Desktop\成都市局第12条.xls'
    result = check_excel(pth)
    print(result)
    # check_excel(r'C:\Users\twj\Desktop\test.xlsx')
