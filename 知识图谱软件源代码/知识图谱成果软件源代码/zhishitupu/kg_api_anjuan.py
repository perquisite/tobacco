import os
import re

import cn2an
import rdflib

from yancaoRegularDemo.Resource.table_19_36.Table34 import *
from yancaoRegularDemo.Resource.tools.utils import get_root_dir
from zhishitupu.src.function import getRootPath
from zhishitupu.tools.discretionary_function import *


def discretionary_power(info_one, info_two, city='成都市局'):
    return_tip = ""
    light_heavy_info_to_print = about_light_heavy(info_two)
    return_tip += light_heavy_info_to_print
    # return light_heavy_info
    for every_case_info in info_one:
        every_tip = ""
        result_dict = grade_judge(every_case_info, city)
        print(result_dict)
        if result_dict == None:
            print(every_case_info['案由'] + " 暂无法处理")
            continue
        if '理应档次' in result_dict.keys() and '实际档次' in result_dict.keys():
            liying = str(result_dict['理应档次'])
            shiji = str(result_dict['实际档次'])

            every_tip += "案由“" + result_dict['案由'] + "”的处罚结果为——"
            for punish in every_case_info['处罚结果']:
                every_tip += "'" + str(punish) + "'" + ","
            every_tip += "属于该地自由裁量基准尺度第" + liying + "项，其基准尺度为第" + shiji + "项，"
            if int(shiji) - int(liying) < 0:
                every_tip += "属于从轻处罚"
            elif int(shiji) - int(liying) > 0:
                every_tip += "属于从重处罚"
        else:
            continue

        if result_dict['罚款'] and result_dict['没收'] and result_dict['收购']:
            pass
        else:
            every_tip += "，且不应该"
        flag1 = True
        if not result_dict['罚款']:
            every_tip += "罚款"
            flag1 = False
        if not result_dict['没收']:
            if flag1:
                every_tip += "没收烟草或烟草制品或违法所得"
                flag1 = False
            else:
                every_tip += "、没收烟草或烟草制品或违法所得"
                flag1 = False
        if not result_dict['收购']:
            if flag1:
                every_tip += "收购"
            else:
                every_tip += "、收购"
        every_tip += "。"
        return_tip += every_tip

    return return_tip


def grade_judge(case_info, city):
    reason = case_info['案由']
    rdfpath = search_4_rdf(reason, city)
    # 获得 阈值单位、惩罚单位
    yz_cf = get_yz_cf_danwei(reason)  # 是个list
    return_info = {'案由': reason}
    if yz_cf:
        print(yz_cf)
        # 去判断档次和从轻从重
        # 根据阈值和惩罚单位的不同，会产生不同的分支
        fact_grade_global = 0
        if yz_cf[0] == '元' and yz_cf[1] in ['比例', '%']:
            """
            {detail_dict: '涉案金额':XXX, '罚款或比例': XXX（罚款总额或罚款比例）,'是否罚款': True False '是否没收': True False,'是否收购':True False}
            涉案金额: 可以是涉案金额、烟草条数、公斤数，只是key名字叫“涉案金额”
            罚款或比例: 可以是罚款金额、罚款比例
            """
            detail_dict = get_detail_m_p(case_info)
            if '涉案金额' in detail_dict.keys():
                print('案件细节:')
                print(detail_dict)
                should_grade = get_should_grade(detail_dict['涉案金额'], rdfpath)  # 获取 应当的 处罚结果的档次
                print('should_grade: ' + str(should_grade))
                fact_grade = get_fact_grade(detail_dict['罚款或比例'], rdfpath)  # 获取 实际的 处罚结果的档次
                fact_grade_global = fact_grade
                print('fact_grade: ' + str(fact_grade))
                return_info['理应档次'] = should_grade
                return_info['实际档次'] = fact_grade

            else:
                return {}

        elif yz_cf[0] == '元' and yz_cf[1] == '元':
            detail_dict = get_detail_m1_m2(case_info)
            if '涉案金额' in detail_dict.keys():
                print('案件细节:')
                print(detail_dict)
                should_grade = get_should_grade(detail_dict['涉案金额'], rdfpath)  # 获取 应当的 处罚结果的档次
                print('should_grade: ' + str(should_grade))
                fact_grade = get_fact_grade(detail_dict['罚款或比例'], rdfpath)  # 获取 实际的 处罚结果的档次
                fact_grade_global = fact_grade
                print('fact_grade: ' + str(fact_grade))
                return_info['理应档次'] = should_grade
                return_info['实际档次'] = fact_grade
            else:
                return {}
        elif yz_cf[0] == '公斤' and yz_cf[1] in ['比例', '%']:
            detail_dict = get_detail_kg_m2(case_info)

        if not fact_grade_global == 0:
            yuyi_yes_or_no = check_yuyi(rdfpath, fact_grade_global)  # 分别是罚款、没收、收购的在rdf里应不应该实行的布尔值list
            print('rdf理应“予以”:')
            print(yuyi_yes_or_no)
            if not yuyi_yes_or_no:
                pass
            # 判断“予以”的处罚对没对
            if '是否罚款' not in detail_dict.keys():
                return_info['罚款'] = True
            else:
                if detail_dict['是否罚款'] == yuyi_yes_or_no[0]:
                    return_info['罚款'] = True
                else:
                    if detail_dict['是否罚款']:
                        return_info['罚款'] = False
                    else:
                        return_info['罚款'] = True

            if '是否没收' not in detail_dict.keys():
                return_info['没收'] = True
            else:
                if detail_dict['是否没收'] == yuyi_yes_or_no[1]:
                    return_info['没收'] = True
                else:
                    if detail_dict['是否没收']:
                        return_info['没收'] = False
                    else:
                        return_info['没收'] = True

            if '是否收购' not in detail_dict.keys():
                return_info['收购'] = True
            else:
                if detail_dict['是否收购'] == yuyi_yes_or_no[2]:
                    return_info['收购'] = True
                else:
                    if detail_dict['是否收购']:
                        return_info['收购'] = False
                    else:
                        return_info['收购'] = True

        return return_info
    else:  # 在rdf没找到阈值单位或惩罚单位
        return None
        print("在rdf没找到阈值单位或惩罚单位，rdf文件出错了")


def get_detail_kg_m2(case_info):
    punishment_lst = case_info['处罚结果']
    detail_dict = {}  # {'涉案金额':XXX, '罚款或比例': XXX（这里是比例）,'是否罚款': True False '是否没收': True False,'是否收购':True False}
    for item in punishment_lst:
        if "处以" in item:  # 提取涉案公斤kg,罚款m2
            m1, p = extract_kg_and_p(item)
            detail_dict['涉案金额'] = m1
            detail_dict['罚款或比例'] = p
            if not p:
                detail_dict['是否罚款'] = False
            else:
                detail_dict['是否罚款'] = True
        elif "予以" in item:
            if "没收" in item:
                detail_dict['是否没收'] = True
            else:
                detail_dict['是否没收'] = False
            if "收购" in item:
                detail_dict['是否收购'] = True
            else:
                detail_dict['是否收购'] = False
    return detail_dict


def extract_kg_and_p(sentence):
    # 提取 钱
    temp = re.search(r"(\d+(\.\d+)?)公斤", sentence)
    if temp:
        kg = float(temp.group(1))
    else:
        kg = None
    # 提取 罚款
    temp = re.search(r"\s*(\d+(\.\d+)?)\s*%", sentence)
    if temp:
        p = float(temp.group(1)) / 100
    else:
        a = cn2an.transform(sentence, "cn2an")
        temp = re.findall(r"\s*(\d+(\.\d+)?)\s*", a)
        if temp:
            p = float(temp[1][0])
        else:
            p = None
    return kg, p

def get_detail_m1_m2(case_info):  # 提出处罚结果中含有 涉案金额 + 罚款金额 的情况,还有“予以没收、收购（罚款）”的True False
    punishment_lst = case_info['处罚结果']
    detail_dict = {}  # {'涉案金额':XXX, '罚款或比例': XXX（这里是比例）,'是否罚款': True False '是否没收': True False,'是否收购':True False}
    for item in punishment_lst:
        if "处以" in item:  # 提取涉案金额m,罚款比例p
            m1, m2 = extract_m1_and_m2(item)
            detail_dict['涉案金额'] = m1
            detail_dict['罚款或比例'] = m2
            if not m2:
                detail_dict['是否罚款'] = False
            else:
                detail_dict['是否罚款'] = True
        elif "予以" in item:
            if "没收" in item:
                detail_dict['是否没收'] = True
            else:
                detail_dict['是否没收'] = False
            if "收购" in item:
                detail_dict['是否收购'] = True
            else:
                detail_dict['是否收购'] = False
    return detail_dict


def extract_m1_and_m2(sentence):   # 默认会写成类似  处以1000.0元涉案金额150元的罚款 的形式 否则失效
    a = cn2an.transform(sentence, "cn2an")
    # 提取 涉案金额
    temp = re.findall(r"(\d+(\.\d+)?)元", a)
    if temp:
        m1 = float(temp[0][0])
        m2 = float(temp[1][0])
    return m1, m2


def get_detail_m_p(case_info):  # 提出处罚结果中含有 涉案金额 + 罚款或比例的情况,还有“予以没收、收购（罚款）”的True False
    punishment_lst = case_info['处罚结果']
    detail_dict = {}  # {'涉案金额':XXX, '罚款或比例': XXX（这里是比例）,'是否罚款': True False '是否没收': True False,'是否收购':True False}
    for item in punishment_lst:
        if "处以" in item:  # 提取涉案金额m,罚款比例p
            m, p = extract_m_and_p(item)
            detail_dict['涉案金额'] = m
            detail_dict['罚款或比例'] = p
            if not p:
                detail_dict['是否罚款'] = False
            else:
                detail_dict['是否罚款'] = True
        elif "予以" in item:
            if "没收" in item:
                detail_dict['是否没收'] = True
            else:
                detail_dict['是否没收'] = False
            if "收购" in item:
                detail_dict['是否收购'] = True
            else:
                detail_dict['是否收购'] = False
    return detail_dict


def extract_m_and_p(sentence):
    # 提取 钱
    temp = re.findall(r"\s*(\d+(\.\d+)?)\s*元", sentence)
    if temp == 2:
        temp.sort()
        m = float(temp[-1][0])
    else:
        m = float(temp[0][0])
    # 提取 比例  （1）XX%  （2）X倍
    temp = re.search(r"\s*(\d*?)\s*%", sentence)
    if temp:
        p = float(temp.group(1)) / 100
    else:
        temp = re.search(r"\s*(([一二三四五六七八九0-9])([\.点][一二三四五六七八九0-9])?)倍\s*", sentence)
        if temp:
            a = temp.group(1)
            p = cn2an.transform(a, "cn2an")
        else:
            temp = re.findall(r"(\d+(\.\d+)?)元", sentence)
            if temp:
                if len(temp) == 2:
                    money = [float(i[0]) for i in temp]
                    money.sort()
                    p = money[0] / money[1]
                else:
                    p = None
            else:
                p = None
    return m, p


def check_yuyi(rdfpath, fact_grade):
    g = rdflib.Graph()
    g.parse(rdfpath, format="xml")
    bool_list = []  # 分别是罚款、没收、收购的应不应该实行的布尔值
    bool_list_origin = []
    q1 = "select ?ys where {<http://www.tobacco.org#处罚结果" + str(fact_grade) + \
         "> <http://www.tobacco.org#罚款> ?ys.}"
    q2 = "select ?ys where {<http://www.tobacco.org#处罚结果" + str(fact_grade) + \
         "> <http://www.tobacco.org#没收烟草或烟草制品> ?ys.}"
    q3 = "select ?ys where {<http://www.tobacco.org#处罚结果" + str(fact_grade) + \
         "> <http://www.tobacco.org#没收违法所得> ?ys.}"
    q4 = "select ?ys where {<http://www.tobacco.org#处罚结果" + str(fact_grade) + \
         "> <http://www.tobacco.org#收购> ?ys.}"
    q = [q1, q2, q3, q4]
    for i in q:
        x = g.query(i)
        t = list(x)  # 二维的
        bool = str(t[0][0]).strip()
        bool_list_origin.append(bool)
    if bool_list_origin[0] == "true":
        bool_list.append(True)
    else:
        bool_list.append(False)
    if bool_list_origin[1] == "true" or bool_list_origin[2] == "true":
        bool_list.append(True)
    else:
        bool_list.append(False)
    if bool_list_origin[3] == "true":
        bool_list.append(True)
    else:
        bool_list.append(False)

    return bool_list


def get_fact_grade(p, rdfpath):   # 若是没有罚款比例只有收购等，则无法判
    g = rdflib.Graph()
    g.parse(rdfpath, format="xml")
    flag = True
    i = 1
    while flag:
        q = "select ?pmin ?pmax where {<http://www.tobacco.org#处罚结果" + str(i) + \
            "> <http://www.tobacco.org#最小惩罚> ?pmin." + \
            "<http://www.tobacco.org#处罚结果" + str(i) + \
            "> <http://www.tobacco.org#最大惩罚> ?pmax.}"
        x = g.query(q)
        # print(x)
        t = list(x)  # 二维的
        if t:
            pmin = float(t[0][0])
            pmax = float(t[0][1])
            if pmin <= p < pmax:
                return i
                # flag = False
            elif pmin == p == pmax:
                return i
                # flag = False
            # elif pmin <= Proportion <= pmax:
            #     flag = False
            else:
                i += 1
        else:
            break
    return i - 1


def get_should_grade(num, rdfpath):
    g = rdflib.Graph()
    g.parse(rdfpath, format="xml")
    q = "select ?yuzhi where {?root <http://www.tobacco.org#阈值> ?yuzhi.}"  # 阈值
    x = g.query(q)
    t = list(x)  # 二维的
    threshold = str(t[0][0]).strip().split()  # threshold是罚款档次（具体金额）
    threshold = list(map(float, threshold))  # threshold是list类型，就算只有一个值
    # 在不考虑从轻/重的情况下，确定理应处罚的档次grade
    grade = 1
    for i in threshold:
        if num >= i:
            grade += 1
        else:
            break
    return grade


def get_yz_cf_danwei(reason, city='成都市局'):
    # 先根据案由，去rdf库里搜索是哪一条
    rdfpath = search_4_rdf(reason, city)
    if rdfpath:
        print("rdfpath: " + rdfpath)
        # 去获得阈值单位、惩罚单位
        g = rdflib.Graph()
        g.parse(rdfpath, format="xml")
        try:
            q = "select ?chengfa where {?root <http://www.tobacco.org#惩罚单位> ?chengfa.}"
            x = g.query(q)
            t = list(x)
            chengfa = t[0][0].strip().split()[0]
            q = "select ?yuzhi where {?root <http://www.tobacco.org#阈值单位> ?yuzhi.}"
            x = g.query(q)
            t = list(x)
            yuzhi = t[0][0].strip().split()[0]
            return [yuzhi, chengfa]
        except Exception as e:
            print("get_yz_cf()函数在查询时产生了一个错误" + str(e.args))
    else:  # 找不到这个案由对应的rdf文件
        return False  # ？？？


def search_4_rdf(reason, city='成都市局'):
    city_folder = getRootPath() + "rdf" + "\\" + city
    # print(city_folder)
    try:
        for f in os.listdir(city_folder):  # listdir返回文件中所有目录
            rdfpath = city_folder + "\\" + f
            g = rdflib.Graph()
            g.parse(rdfpath, format="xml")
            q = "select ?reason where {?root <http://www.tobacco.org#违法行为名称> ?reason.}"
            # q = "select ?threshold where {?root <http://www.tobacco.org#阈值> ?threshold.}"  # 阈值
            x = g.query(q)
            t = list(x)
            r = str(t[0][0]).strip().split()[0]
            if reason == r:
                return rdfpath
        return False
    except Exception as e:
        return "search_4_rdf出现了一个错误" + str(e.args)


def about_light_heavy(info_two):
    if info_two['从轻']:
        light_info = "“承办人意见”中存在“从轻“字样，"
        if info_two['存在从轻证据']:
            light_info += "且存在从轻证据。"
        else:
            light_info += "但不存在从轻证据！"
    else:
        light_info = "“承办人意见”中不存在“从轻“字样，"
        if info_two['存在从轻证据']:
            light_info += "但存在从轻证据。"
        else:
            light_info += "且不存在从轻证据。"
    if info_two['从重']:
        heavy_info = "“承办人意见”中存在“从重“字样。"
    else:
        heavy_info = "“承办人意见”中不存在“从重“字样。"
    light_heavy_info = light_info + heavy_info
    return light_heavy_info


if __name__ == '__main__':
    # t = Table34("C:\\Users\\twj\Desktop\\1\\", "C:\\Users\\twj\Desktop\\1\\"))
    # print(get_yz_cf('未在当地烟草专卖批发企业进货', '成都市局'))
    # case_info = [{'案由': '未在当地烟草专卖批发企业进货', '处罚结果': ['处以1001元 6%的罚款。', '予以没收、收购。']},]
    #              {'案由': '销售、运输、存储、投递走私烟草制品，出口倒流国产烟草制品、未缴付关税而流出免税店和保税区的烟草制品', '处罚结果': ['处以50002元涉案金额2001元的罚款。', '予以收购。']}]
    # grade_judge(case_info, city='成都市局')
    case_info = [{'案由': '销售、运输、存储、投递走私烟草制品，出口倒流国产烟草制品、未缴付关税而流出免税店和保税区的烟草制品', '处罚结果': ['处以60001元 3000元的罚款。']}]
    return_info_two = {'从轻': True, '存在从轻证据': ['xxxxxx', 'yyyy'], '从重': False}
    print(discretionary_power(case_info, return_info_two))

    # info_one = ""
    # info_two = {'从轻': True, '存在从轻证据': True, '从重': True}
    # discretionary_power(info_one, info_two)
