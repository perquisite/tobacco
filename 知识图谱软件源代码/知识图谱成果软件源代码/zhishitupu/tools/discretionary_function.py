import os
import re

import cn2an
import jieba
import rdflib

from zhishitupu.src.function import getRootPath


def get_every_detail(written_reason, real_reason, punishment, city):
    temp_list = {'案由': written_reason}
    rdfpath = search_4_rdf(real_reason, city)
    print('rdfpath: ')
    print(rdfpath)
    yz_cf_danwei = get_yz_cf_danwei(real_reason, city)
    # print('阈值和惩罚单位: ')
    # print(yz_cf_danwei)
    if yz_cf_danwei:
        # temp_dict：{理应当次：None/数字 ，实际档次：None/数字，应该没收：T/F， 实际没收：T/F，是否罚款：T/F，理应条文：‘’}
        temp_dict = handle_punishment(punishment, yz_cf_danwei, rdfpath)
        info_processed = dict(temp_list, **temp_dict)
        # 合并info_processed和temp_dict
        return info_processed
    else:  # 没找到单位
        print("在rdf没找到阈值单位或惩罚单位，rdf文件出错了")
        return None


def handle_punishment(punishment, yz_cf_danwei, rdfpath):
    # temp_dict：{理应当次：数/None , 实际档次：数/None, 应该没收：T/F， 实际没收：T/F，是否罚款：T/F，理应条文：‘’}
    temp_dict = {}
    if yz_cf_danwei[0] == '元' and yz_cf_danwei[1] in ['比例', '%']:
        temp_dict = get_temp_dict_m_p(punishment, rdfpath)
        return temp_dict
    elif yz_cf_danwei[0] == '元' and yz_cf_danwei[1] == '元':
        temp_dict = get_temp_dict_m1_m2(punishment, rdfpath)
        return temp_dict
    elif yz_cf_danwei[0] == '元' and yz_cf_danwei[1] == '倍':
        temp_dict = get_temp_dict_m1_times(punishment, rdfpath)
        return temp_dict
    elif yz_cf_danwei[0] == '条' and yz_cf_danwei[1] in ['比例', '%']:
        temp_dict = get_temp_dict_tiao_p(punishment, rdfpath)
        return temp_dict
    elif yz_cf_danwei[0] == '公斤' and yz_cf_danwei[1] in ['比例', '%']:
        temp_dict = get_temp_dict_gongjin_p(punishment, rdfpath)
        return temp_dict
    else:
        pass


def get_temp_dict_gongjin_p(punishment, rdfpath):
    # temp_dict：{理应档次：数 / None, 实际档次：数 / None, 应该没收：T / F， 实际没收：T / F，是否罚款：T / F，理应条文：‘’}
    temp_dict = {}
    temp_dict['是否罚款'] = False
    fakuan_detected = False

    temp_dict['理应档次'] = None
    temp_dict['实际档次'] = None
    for p in punishment:
        if '罚款' in p:
            fakuan_detected = True
            temp_dict['是否罚款'] = True
            a = cn2an.transform(p, "cn2an")
            temp = re.search(r"\s*(\d+(\.\d+)?)\s*公斤", a)
            if temp:
                temp = [float(i[0]) for i in temp]
                temp.sort(reverse=True)  # 从大到小排序，取最大的当条数
                tiao = temp[0]
                temp_dict['理应档次'] = get_should_grade(tiao, rdfpath)
                temp_dict['理应条文'] = get_tiaowen(rdfpath, temp_dict['理应档次'])
            else:
                temp_dict['理应档次'] = None
            temp = re.search(r"\s*(\d+(\.\d+)?)\s*%", a)
            if temp:
                p = float(temp.group(1)) / 100
                temp_dict['实际档次'] = get_fact_grade(p, rdfpath)
            else:
                temp_dict['实际档次'] = None
    # 判断没收可是否用来做理应档次和实际档次的确定
    if not fakuan_detected:
        # if temp_dict['理应档次'] is None or temp_dict['实际档次'] is None:
        temp_dict['实际没收'] = False
        for p in punishment:
            if '没收' in p or '销毁' in p:
                temp_dict['实际没收'] = True
                a = cn2an.transform(p, "cn2an")
                temp = re.findall(r"\s*(\d+(\.\d+)?)\s*公斤", a)
                if temp:
                    temp = [float(i[0]) for i in temp]
                    temp.sort(reverse=True)  # 从大到小排序，取最大的当条数
                    tiao = temp[0]
                    temp_dict['理应档次'] = get_should_grade(tiao, rdfpath)
                    temp_dict['理应条文'] = get_tiaowen(rdfpath, temp_dict['理应档次'])
                else:
                    temp_dict['理应档次'] = None
                # 遍历rdf，确定哪一档独含“没收”，以此确定 实际档次
                bool_list = get_moshou_dangci(rdfpath)
                # print(bool_list)
                ture_index_list = [i for i, j in enumerate(bool_list) if j is True]
                if len(ture_index_list) == 1:
                    temp_dict['实际档次'] = ture_index_list[0] + 1
                else:
                    temp_dict['实际档次'] = None

    temp_dict['应该没收'] = False
    if not temp_dict['理应档次'] is None:  # 理应当次 有
        # 判断在RDF里“没收”是否应该
        temp_dict['应该没收'] = check_moshou(rdfpath, temp_dict['理应档次'])

    return temp_dict


def get_temp_dict_tiao_p(punishment, rdfpath):  # 条 + 比例， 特殊判断没收可是否用来做理应档次和实际档次的确定
    # temp_dict：{理应档次：数 / None, 实际档次：数 / None, 应该没收：T / F， 实际没收：T / F，是否罚款：T / F，理应条文：‘’}
    temp_dict = {}
    temp_dict['是否罚款'] = False
    fakuan_detected = False

    temp_dict['理应档次'] = None
    temp_dict['实际档次'] = None
    for p in punishment:
        if '罚款' in p:
            fakuan_detected = True
            temp_dict['是否罚款'] = True
            a = cn2an.transform(p, "cn2an")
            temp = re.search(r"\s*(\d+(\.\d+)?)\s*条", a)
            if temp:
                temp = [float(i[0]) for i in temp]
                temp.sort(reverse=True)  # 从大到小排序，取最大的当条数
                tiao = temp[0]
                temp_dict['理应档次'] = get_should_grade(tiao, rdfpath)
                temp_dict['理应条文'] = get_tiaowen(rdfpath, temp_dict['理应档次'])
            else:
                temp_dict['理应档次'] = None
            temp = re.search(r"\s*(\d+(\.\d+)?)\s*%", a)
            if temp:
                p = float(temp.group(1)) / 100
                temp_dict['实际档次'] = get_fact_grade(p, rdfpath)
            else:
                temp_dict['实际档次'] = None
    # 判断没收可是否用来做理应档次和实际档次的确定
    if not fakuan_detected:
    # if temp_dict['理应档次'] is None or temp_dict['实际档次'] is None:
        temp_dict['实际没收'] = False
        for p in punishment:
            if '没收' in p or '销毁' in p:
                temp_dict['实际没收'] = True
                a = cn2an.transform(p, "cn2an")
                temp = re.findall(r"\s*(\d+(\.\d+)?)\s*条", a)
                if temp:
                    temp = [float(i[0]) for i in temp]
                    temp.sort(reverse=True)  # 从大到小排序，取最大的当条数
                    tiao = temp[0]
                    temp_dict['理应档次'] = get_should_grade(tiao, rdfpath)
                    temp_dict['理应条文'] = get_tiaowen(rdfpath, temp_dict['理应档次'])
                else:
                    temp_dict['理应档次'] = None
                # 遍历rdf，确定哪一档独含“没收”，以此确定 实际档次
                bool_list = get_moshou_dangci(rdfpath)
                # print(bool_list)
                ture_index_list = [i for i, j in enumerate(bool_list) if j is True]
                if len(ture_index_list) == 1:
                    temp_dict['实际档次'] = ture_index_list[0] + 1
                else:
                    temp_dict['实际档次'] = None

    temp_dict['应该没收'] = False
    if not temp_dict['理应档次'] is None:  # 理应当次 有
        # 判断在RDF里“没收”是否应该
        temp_dict['应该没收'] = check_moshou(rdfpath, temp_dict['理应档次'])

    return temp_dict


def get_temp_dict_m1_times(punishment, rdfpath):  # 元 + 倍
    # temp_dict：{理应档次：数 / None, 实际档次：数 / None, 应该没收：T / F， 实际没收：T / F，是否罚款：T / F，理应条文：‘’}
    temp_dict = {}
    temp_dict['是否罚款'] = False

    temp_dict['理应档次'] = None
    temp_dict['实际档次'] = None
    for p in punishment:
        if '罚款' in p:
            temp_dict['是否罚款'] = True
            a = cn2an.transform(p, "cn2an")
            temp = re.search(r"\s*(\d+(\.\d+)?)\s*元", a)  # 可能会有数字大于两个的情况？
            if temp:
                m = float(temp.group(1))
                temp_dict['理应档次'] = get_should_grade(m, rdfpath)
                temp_dict['理应条文'] = get_tiaowen(rdfpath, temp_dict['理应档次'])
                # break
            else:
                temp = re.search(r"\s*(\d+(\.\d+)?)\s*倍", a)
                if temp:
                    p = float(temp.group(1))
                    temp_dict['实际档次'] = get_fact_grade(p, rdfpath)
                else:
                    temp_dict['实际档次'] = None

    temp_dict['应该没收'] = False
    temp_dict['实际没收'] = False
    if (not temp_dict['理应档次'] is None) and (not temp_dict['实际档次'] is None):  # 理应当次 和 实际当次 都有
        # 判断在RDF里“没收”是否应该
        temp_dict['应该没收'] = check_moshou(rdfpath, temp_dict['实际档次'])
        # 判断在处理结果里有无“没收”
        temp_dict['实际没收'] = False
        for p in punishment:
            if '没收' in p or '销毁' in p:
                temp_dict['实际没收'] = True
                break

    return temp_dict


def get_temp_dict_m1_m2(punishment, rdfpath):
    # temp_dict：{理应档次：数 / None, 实际档次：数 / None, 应该没收：T / F， 实际没收：T / F，是否罚款：T / F，理应条文：‘’}
    temp_dict = {}
    temp_dict['是否罚款'] = False

    temp_dict['理应档次'] = None
    temp_dict['实际档次'] = None
    for p in punishment:
        if '罚款' in p:
            temp_dict['是否罚款'] = True
            a = cn2an.transform(p, "cn2an")
            temp = re.findall(r"\s*(\d+(\.\d+)?)\s*元", a)  # 可能会有数字大于两个的情况？
            if temp:
                temp = [float(i[0]) for i in temp]
                # temp.sort(key=None, reverse=True) 不能排序，因为罚款不一定小于涉案金额，可能大于
                m1 = temp[0]
                m2 = temp[1]
                temp_dict['理应档次'] = get_should_grade(m1, rdfpath)
                temp_dict['理应条文'] = get_tiaowen(rdfpath, temp_dict['理应档次'])
                temp_dict['实际档次'] = get_fact_grade(m2, rdfpath)
                # break
            else:
                temp_dict['理应档次'] = None
                temp_dict['实际档次'] = None

    temp_dict['应该没收'] = False
    temp_dict['实际没收'] = False
    if (not temp_dict['理应档次'] is None) and (not temp_dict['实际档次'] is None):  # 理应当次 和 实际当次 都有
        # 判断在RDF里“没收”是否应该
        temp_dict['应该没收'] = check_moshou(rdfpath, temp_dict['实际档次'])
        # 判断在处理结果里有无“没收”
        temp_dict['实际没收'] = False
        for p in punishment:
            if '没收' in p or '销毁' in p:
                temp_dict['实际没收'] = True
                break

    return temp_dict


def get_temp_dict_m_p(punishment, rdfpath):
    # temp_dict：{理应档次：数 / None, 实际档次：数 / None, 应该没收：T / F， 实际没收：T / F，是否罚款：T / F，理应条文：‘’}
    temp_dict = {}
    temp_dict['是否罚款'] = False
    # fakuan_detected = False

    temp_dict['理应档次'] = None
    temp_dict['实际档次'] = None
    for p in punishment:
        if '罚款' in p:
            # fakuan_detected = True
            temp_dict['是否罚款'] = True
            a = cn2an.transform(p, "cn2an")
            temp = re.search(r"\s*(\d+(\.\d+)?)\s*元", a)
            if temp:
                m = float(temp.group(1))
                temp_dict['理应档次'] = get_should_grade(m, rdfpath)
                temp_dict['理应条文'] = get_tiaowen(rdfpath, temp_dict['理应档次'])
            else:
                temp_dict['理应档次'] = None
            temp = re.search(r"\s*(\d+(\.\d+)?)\s*%", a)
            if temp:
                p = float(temp.group(1)) / 100
                temp_dict['实际档次'] = get_fact_grade(p, rdfpath)
            else:
                temp_dict['实际档次'] = None

    temp_dict['应该没收'] = False
    temp_dict['实际没收'] = False
    if (not temp_dict['理应档次'] is None) and (not temp_dict['实际档次'] is None):  # 理应当次 和 实际当次 都有
        # 判断在RDF里“没收”是否应该
        temp_dict['应该没收'] = check_moshou(rdfpath, temp_dict['实际档次'])
        # 判断在处理结果里有无“没收”
        temp_dict['实际没收'] = False
        for p in punishment:
            if '没收' in p or '销毁' in p:
                temp_dict['实际没收'] = True
                break

    return temp_dict
    # elif (not temp_dict['理应档次'] is None) and temp_dict['实际档次'] is None: # 理应当次有 实际当次无
    # 借助“没收”确定实际档次


def get_tiaowen(rdfpath, should_grade):
    g = rdflib.Graph()
    g.parse(rdfpath, format="xml")
    q = "select ?fatiao where {?root <http://www.tobacco.org#法律条数> ?fatiao.}"
    x = g.query(q)
    t = list(x)
    fatiao = cn2an.transform(t[0][0].strip().split()[0], "cn2an")
    q1 = "select ?ys where {<http://www.tobacco.org#违法行为" + str(fatiao) + "处罚结果" + str(should_grade) + \
         "> <http://www.tobacco.org#处罚结果原文> ?ys.}"
    x = g.query(q1)
    t = list(x)
    tiaowen = str(t[0][0]).strip()
    return tiaowen


def check_moshou(rdfpath, fact_grade):
    g = rdflib.Graph()
    g.parse(rdfpath, format="xml")
    q = "select ?fatiao where {?root <http://www.tobacco.org#法律条数> ?fatiao.}"
    x = g.query(q)
    t = list(x)
    fatiao = cn2an.transform(t[0][0].strip().split()[0], "cn2an")
    q1 = "select ?ys where {<http://www.tobacco.org#违法行为" + str(fatiao) + "处罚结果" + str(fact_grade) + \
         "> <http://www.tobacco.org#没收烟草或烟草制品> ?ys.}"
    q2 = "select ?ys where {<http://www.tobacco.org#违法行为" + str(fatiao) + "处罚结果" + str(fact_grade) + \
         "> <http://www.tobacco.org#没收违法所得> ?ys.}"
    x = g.query(q1)
    t = list(x)  # 二维的
    q1_result = str(t[0][0]).strip()
    x = g.query(q2)
    t = list(x)  # 二维的
    q2_result = str(t[0][0]).strip()
    # print(q1_result, q2_result)
    if q1_result == 'true' or q2_result == 'true':
        return True
    else:
        return False


def get_fact_grade(p, rdfpath):  # 若是没有罚款比例只有收购等，则无法判
    g = rdflib.Graph()
    g.parse(rdfpath, format="xml")
    q1 = "select ?fatiao where {?root <http://www.tobacco.org#法律条数> ?fatiao.}"
    x = g.query(q1)
    t = list(x)
    fatiao = cn2an.transform(t[0][0].strip().split()[0], "cn2an")
    flag = True
    i = 1
    while flag:
        q = "select ?pmin ?pmax where {<http://www.tobacco.org#违法行为" + str(fatiao) + "处罚结果" + str(i) + \
            "> <http://www.tobacco.org#最小惩罚> ?pmin. " + \
            "<http://www.tobacco.org#违法行为" + str(fatiao) + "处罚结果" + str(i) + \
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
    # # 先确定共几档
    # q = "select ?yuzhi where {?root <http://www.tobacco.org#处罚> ?yuzhi.}"  # 阈值
    # x = g.query(q)
    # t = list(x)  # 二维的
    # dangci_num = len(t)
    # # 确定发条数
    # q = "select ?fatiao where {?root <http://www.tobacco.org#法律条数> ?fatiao.}"
    # x = g.query(q)
    # t = list(x)
    # fatiao = cn2an.transform(t[0][0].strip().split()[0], "cn2an")
    # i = 1
    # while i <= dangci_num:
    #     q1 = "select ?ys where {<http://www.tobacco.org#违法行为" + str(fatiao) + "处罚结果" + str(i) + \
    #          "> <http://www.tobacco.org#最小惩罚> ?ys.}"
    #     q2 = "select ?ys where {<http://www.tobacco.org#违法行为" + str(fatiao) + "处罚结果" + str(i) + \
    #          "> <http://www.tobacco.org#最大惩罚> ?ys.}"
    #     x = g.query(q1)
    #     t = list(x)  # 二维的
    #     p_min = int(str(t[0][0]).strip())
    #     x = g.query(q2)
    #     t = list(x)  # 二维的
    #     p_max = int(str(t[0][0]).strip())
    #     print(p_min, p_max)
    #     if p_min <= num < p_max:
    #         return i
    #
    #     i += 1


def get_moshou_dangci(rdfpath):
    g = rdflib.Graph()
    g.parse(rdfpath, format="xml")
    # 先确定共几档
    q = "select ?yuzhi where {?root <http://www.tobacco.org#处罚> ?yuzhi.}"  # 阈值
    x = g.query(q)
    t = list(x)  # 二维的
    dangci_num = len(t)
    # 确定发条数
    q = "select ?fatiao where {?root <http://www.tobacco.org#法律条数> ?fatiao.}"
    x = g.query(q)
    t = list(x)
    fatiao = cn2an.transform(t[0][0].strip().split()[0], "cn2an")
    i = 1
    bool_list = []
    while i <= dangci_num:
        q1 = "select ?ys where {<http://www.tobacco.org#违法行为" + str(fatiao) + "处罚结果" + str(i) + \
             "> <http://www.tobacco.org#没收烟草或烟草制品> ?ys.}"
        q2 = "select ?ys where {<http://www.tobacco.org#违法行为" + str(fatiao) + "处罚结果" + str(i) + \
             "> <http://www.tobacco.org#没收违法所得> ?ys.}"
        x = g.query(q1)
        t = list(x)  # 二维的
        bool1 = str(t[0][0]).strip()
        x = g.query(q2)
        t = list(x)  # 二维的
        bool2 = str(t[0][0]).strip()
        print(bool1, bool2)
        if bool1 == 'true' or bool2 == 'true':
            bool_list.append(True)
        else:
            bool_list.append(False)

        i += 1
    return bool_list


def get_yz_cf_danwei(reason, city):
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


def getRealReason(reason):
    reason_list = ['无烟草专卖品准运证运输烟草专卖品', '无烟草专卖品准运证运输烟草专卖品（承运人）',
                   "邮寄、异地携带烟叶、烟草制品超过国务院有关部门规定的限量",
                   '无烟草专卖生产企业许可证生产烟草制品', '无烟草专卖生产企业许可证生产卷烟纸、滤嘴棒、烟用丝束或烟草专用机械',
                   '无烟草专卖批发企业许可证经营烟草制品批发业务', '销售非法生产的烟草专卖品', '使用涂改、伪造、变造的烟草专卖许可证',
                   '不及时办理烟草专卖许可证变更、注销手续', '烟草专卖零售经营者向未成年人出售烟草制品或未亮证经营',
                   '无烟草专卖零售许可证经营烟草制品零售业务', '未在当地烟草专卖批发企业进货',
                   '生产、销售、运输、存储、投递假冒伪劣烟草制品',
                   '为生产、销售、运输、存储、投递假冒伪劣烟草制品提供条件',
                   '擅自收购烟叶',
                   '擅自跨省、自治区、直辖市从事烟草制品批发业务', '为无烟草专卖许可证的单位或者个人提供烟草专卖品',
                   '从无烟草专卖生产企业许可证、特种烟草专卖经营企业许可证的企业购买卷烟纸、滤嘴棒、烟用丝束、烟草专用机械',
                   '免税进口的烟草制品不按规定存放在烟草制品保税仓库内',
                   '在海关监管区内经营免税的卷烟、雪茄烟没有在小包、条包上标注国务院烟草专卖行政主管部门规定的专门标志',
                   '拍卖企业未对竞买人进行资格验证，或者不接受烟草专卖行政主管部门的监督，擅自拍卖烟草专卖品',
                   '申请人隐瞒有关情况或者提供虚假材料',
                   '销售、运输、存储、投递走私烟草制品',
                   '为销售、运输、存储、投递走私烟草制品提供条件',
                   '超越经营范围和地域范围从事烟草制品批发业务',
                   '向未成年人出售烟草制品被烟草专卖行政主管部门责令改正，但拒不改正',
                   '未亮证经营被烟草专卖行政主管部门责令改正，但拒不改正', '向未成年人销售卷烟',
                   '向未成年人销售电子烟', '烟草专卖零售经营者拒绝接受检查']
    seg = jieba.lcut(reason)
    # print(seg)
    index_j = None
    # index_i = None
    for i in reason_list:
        index_i = reason_list.index(i)
        for j in seg:
            if j in i:
                index_j = seg.index(j)
                continue
            else:
                break
        if index_j == len(seg) - 1:
            return i
        else:
            if not index_i == len(reason_list) - 1:
                continue

    return None


if __name__ == "__main__":
    # print(get_should_grade(250, r'D:\twj_new\tobacco\zhishitupu\rdf\成都市局\7.rdf'))
    real_reason = getRealReason('存储假冒伪劣烟草制品')
    print(real_reason)
    print(search_4_rdf(real_reason))
