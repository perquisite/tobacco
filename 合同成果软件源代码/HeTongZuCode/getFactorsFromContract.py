# -*- coding:utf-8 -*-
# @ModuleName: getFactorsFromContract
# @Function: 
# @Author: huhonghui
# @email: 1241328737@qq.com
# @Time: 2021/6/18 19:56
import os
import re

import pythoncom

from architecture_contract import processFunc3
from utils import UnifiedSocialCreditIdentifier, checkIdCard, isTelPhoneNumber, isRightDate, checkQQ, checkEmail, \
    str_insert, addRemarkInDoc, digital_to_Upper, is_contain_dot, isEmail, check_str
from win32com.client import Dispatch
import helpful as hp
from purchase_and_warehousing_contract import processFunc
import None_standard_contract
import datetime
from rent_contract import processFuncRent


# buy_sell_contract
# Modified by WZK
# 批注功能实现
def buy_sell_contract(text, filePath, processed_file_sava_dir):
    # print('进入getFactorsFromContract.py下的buy_sell_contract')
    # print(text)
    flag = [False]
    goods_flag_1 = False
    goods_flag_2 = False
    goods_flag_3 = False
    goods_flag_4 = False
    goods_flag_5 = False
    goods_flag_6 = False
    goods_flag_7 = False

    matrial_flag_1 = False
    matrial_flag_2 = False

    contractExecution_flag_1 = False
    contractExecution_flag_2 = False
    contractExecution_flag_3 = False
    contractExecution_flag_4 = False
    contractExecution_flag_5 = False
    contractExecution_flag_6 = False

    check_flag_1 = False
    check_flag_2 = False
    check_flag_3 = False

    account_flag_1 = False
    account_flag_2 = False
    account_flag_3 = False
    account_flag_4 = False
    account_flag_5 = False
    account_flag_6 = False

    account1_flag_1 = False
    account1_flag_2 = False
    account1_flag_3 = False
    account1_flag_4 = False
    account1_flag_5 = False
    account1_flag_6 = False

    feedback_flag_1 = False
    feedback_flag_2 = False
    feedback_flag_3 = False
    feedback_flag_4 = False
    feedback_flag_5 = False

    try:
        pythoncom.CoInitialize()
        word = Dispatch('Word.Application')
        pythoncom.CoInitialize()
        word.Documents.close()
        word.Quit()
        word.Visible = 0  # 后台打开word文档
    except Exception as ex:
        print(ex)
    try:
        document = word.Documents.Open(FileName=filePath)
    except Exception as ex:
        print(ex)
    factors = {}
    factors_ok = []
    factors_error = {}
    factors_to_inform = {}
    factors_miss = []
    factors_miss_block = ["智能合同审查结果："]

    # 下面开始通过text文本来匹配要审查的要素
    # 例子：
    # text中的“甲方（买方）：xxx” 可以使用正则表达式 match = '甲方（买方）：(.*?)\n'来匹配出“xxx”  re.findall
    # 存储 factors["甲方（买方）"] = “xxx”
    # 与此同时可以判断这个要素是否ok，如果ok: factors_ok.append("甲方（买方）") ## 若要素的名称后文再次出现，需要区分
    # 若要素不ok， factors_error["甲方（买方）"] = “错误原因”
    # 对于系统需要提示给业务人员现场核查确认比对的信息，添加至factors_to_inform：factors_to_inform[位置信息相关（用于定位）] = 提示内容
    # print(text.count("联系方式"))
    # 0.甲方乙方主题审查
    # 甲方主体审查
    # 甲方名称审查
    # print("甲方（买方）：" in text)

    # flag全为1，就不同批注大项缺失
    # hp.addRemarkInDoc(word, document, "\n", "智能合同审查结果：\n")
    flag = 1

    try:
        match = '甲方（买方）：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")  # 找到主体名并去掉空格
        factors["甲方（买方）"] = factor
        if factor != "":
            factors_ok.append("甲方（买方）")
        else:
            factors_error["甲方（买方）"] = "未填写甲方姓名"
            hp.addRemarkInDoc(word, document, "甲方（买方）", "甲方（买方）主体信息不完整")
            # factors_to_inform['甲方（买方）审核提示'] = "请审核甲方（买方）是否填写"
    except:
        factors_miss.append("甲方（买方）合同要素缺失")
        flag = 0
        pass

    # 居住地审查
    # 新方法，测试
    try:
        match = '住所地：(.*?)\n'
        factor_location = re.findall(match, text)[0].replace(" ", "")
        # print(factor_location)
        factors["甲方（买方）住所地"] = factor_location
        if factor_location != '':
            factors_ok.append("甲方（买方）住所地")
        else:
            factors_error["甲方住所地"] = "甲方住所地未填写完整"
            hp.addRemarkInDoc(word, document, "住所地", "甲方住所地不完整")
            # factors_to_inform['甲方（买方）住所地审核提示'] = "请审核甲方（买方）住所地是否填写"
    except:
        factors_miss.append("甲方住所地关合同要素缺失")
        flag = 0
        pass

    # 法定代表人审查
    # 一共有五个法定代表人要素，暂时先不审查
    try:
        match = '法定代表人/负责人：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["甲方（买方）法定代表人/负责人"] = factor
        if factor != "":
            factors_ok.append("甲方（买方）法定代表人/负责人")
        else:
            factors_error["甲方（买方）法定代表人/负责人"] = "甲方法定代表人/负责人未填写完整"
            hp.addRemarkInDoc(word, document, "法定代表人/负责人", "甲方（买方）法定代表人/负责人不完整")
            # factors_to_inform['甲方（买方）法定代表人审核提示'] = "请审核甲方（买方）法定代表人是否填写"
    except:
        factors_miss.append("甲方（买方）法定代表人合同要素缺失")
        flag = 0

    # 统一社会信用代码/身份证号码
    # 共两个要素
    # 第一个版本，有错误
    try:
        match = '统一社会信用代码/身份证号码：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["甲方（买方）统一社会信用代码/身份证号码"] = factor
        # print(factor)
        if factor != "":
            if UnifiedSocialCreditIdentifier().check_code(factor, 'sc'):
                factors_ok.append("甲方（买方）统一社会信用代码/身份证号码")
            else:

                if checkIdCard(factor) == 'ok':
                    factors_ok.append("甲方（买方）统一社会信用代码/身份证号码")
                else:
                    factors_error["甲方（买方）统一社会信用代码/身份证号码"] = "统一社会信用代码未填写正确或" + checkIdCard(factor)
                    hp.addRemarkInDoc(word, document, "统一社会信用代码/身份证号码", "甲方统一社会信用代码/身份证号码校验未通过")
                    # factors_to_inform['甲方（买方）统一社会信用代码/身份证号码审核提示'] = "请审核甲方（买方）统一社会信用代码/身份证号码是否填写正确"
        else:
            factors_error["甲方（买方）统一社会信用代码/身份证号码"] = "统一社会信用代码/身份证号码未填写完整"
            # factors_to_inform['甲方（买方）统一社会信用代码/身份证号码审核提示'] = "请审核甲方（买方）统一社会信用代码/身份证号码是否填写正确"
            hp.addRemarkInDoc(word, document, "统一社会信用代码/身份证号码", "甲方统一社会信用代码/身份证号码未填写")
    except:
        factors_miss.append("甲方（买方）统一社会代码/身份证号码合同要素缺失")
        flag = 0
        pass

    # 联系方式，不判断是否正确
    try:
        match = '联系方式：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["甲方（买方）联系方式"] = factor
        if factor != "":
            factor = factor.replace("（", "").replace("）", "").replace("-", "")
            # if len(factor)==10:
            # factors_ok.append("甲方（买方）联系方式")
            # hp.addRemarkInDoc(word, document, "联系方式", "联系方式为10位号码，请人工检查")
            # else:
            if isTelPhoneNumber(factor) != "Error":
                factors_ok.append("甲方（买方）联系方式")
            else:
                factors_error["甲方（买方）联系方式"] = "联系方式填写有误"
                # factors_to_inform['甲方（买方）联系方式审核提示'] = "请审核甲方（买方）联系方式是否填写正确"
                hp.addRemarkInDoc(word, document, "联系方式", "甲方（买方）联系方式错误")
        else:
            factors_error["甲方（买方）联系方式"] = "甲方（买方）联系方式不完整"
            # factors_to_inform['甲方（买方）联系方式审核提示'] = "请审核甲方（买方）联系方式是否填写正确"
            hp.addRemarkInDoc(word, document, "联系方式", "甲方（买方）联系方式未填写完整")
    except:
        factors_miss.append("甲方（买方）联系方式合同要素缺失")
        flag = 0
        pass

    # 邮箱
    try:
        match = '电子邮箱：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["甲方（买方）电子邮箱"] = factor
        # print(factor)
        if factor != "":
            if checkEmail(factor):
                # if check_str("\w{0,19}@[0-9a-zA-Z]{1,13}\.[com,cn,net]{1,3}",factor):
                # print("判断错误")
                factors_ok.append("甲方（买方）电子邮箱")
            else:
                factors_error["甲方（买方）电子邮箱"] = "甲方（买方）电子邮箱填写有误"
                hp.addRemarkInDoc(word, document, "电子邮箱", "甲方（买方）电子邮箱错误")
                # factors_to_inform['甲方（买方）电子邮箱'] = "甲方（买方）电子邮箱填写有误"
        else:
            factors_error["甲方（买方）电子邮箱"] = "甲方（买方）电子邮箱未填写完整"
            hp.addRemarkInDoc(word, document, "电子邮箱", "甲方（买方）电子邮箱填写不完整")
            # factors_to_inform['甲方（买方）电子邮箱'] = "甲方（买方）电子邮箱填写有误"
    except:
        factors_miss.append("甲方（买方）电子邮箱合同要素缺失")
        flag = 0

    # print(factors, factors_ok, factors_error, factors_to_inform)
    # ----------甲方检查完毕--------------#

    # 乙方（卖方）检查
    try:
        match = '乙方（卖方）：(.*?)\n'  # 此处冒号为英文冒号，特此标注
        factor = re.findall(match, text)[0].replace(" ", "")  # 找到主体名并去掉空格
        # print(factor)
        factors["乙方（卖方）"] = factor
        if factor != "":
            factors_ok.append("乙方（卖方）")
        else:
            factors_error["乙方（卖方）"] = "未填写乙方（卖方）姓名"
            hp.addRemarkInDoc(word, document, "乙方（卖方）", "乙方（卖方）主体信息不完整")
            # factors_to_inform['乙方（卖方）审核提示'] = "请审核乙方（卖方）是否填写"
    except:
        factors_miss.append("乙方（卖方）合同要素缺失")
        flag = 0
        pass
    # print(factors, factors_ok, factors_error, factors_to_inform)

    # 居住地审查

    try:
        match = '住所地：(.*?)\n'
        factor = re.findall(match, text)[1].replace(" ", "")
        factors["乙方（卖方）住所地"] = factor
        if factor != '':
            factors_ok.append("乙方（卖方）住所地")
        else:
            factors_error["乙方（卖方）住所地"] = "甲方住所地未填写完整"
            hp.addRemarkInDoc(word, document, "乙方（卖方）", "乙方住所地不完整")
            # factors_to_inform['乙方（卖方）住所地审核提示'] = "请审核乙方（卖方）住所地是否填写"
    except:
        factors_miss.append("乙方住所地合同要素缺失")
        flag = 0
        pass

    # 法定代表人审查
    try:
        match = '法定代表人/负责人：(.*?)\n'
        factor = re.findall(match, text)[1].replace(" ", "")
        factors["乙方（卖方）法定代表人/负责人"] = factor
        if factor != "":
            factors_ok.append("乙方（卖方）法定代表人/负责人")
        else:
            factors_error["乙方（卖方）法定代表人/负责人"] = "乙方（卖方）法定代表人/负责人未填写完整"
            hp.addRemarkInDoc(word, document, "乙方（卖方）", "乙方（卖方）法定代表人/负责人不完整")
            # factors_to_inform['乙方（卖方）法定代表人审核提示'] = "请审核乙方（卖方）法定代表人是否填写"
    except:
        factors_miss.append("乙方（卖方）法定代表人合同要素缺失")
        flag = 0

    # 统一社会信用代码/身份证号码
    try:
        match = '统一社会信用代码/身份证号码：(.*?)\n'
        factor = re.findall(match, text)[1].replace(" ", "")
        factors["乙方（卖方）统一社会信用代码/身份证号码"] = factor
        if factor != "":
            if UnifiedSocialCreditIdentifier().check_code(factor, 'sc'):
                factors_ok.append("乙方（卖方）统一社会信用代码/身份证号码")
            else:

                if checkIdCard(factor) == 'ok':
                    factors_ok.append("乙方（卖方）统一社会信用代码/身份证号码")
                else:
                    factors_error["乙方（卖方）统一社会信用代码/身份证号码"] = "统一社会信用代码未填写正确或" + checkIdCard(factor)
                    hp.addRemarkInDoc(word, document, "乙方（卖方）", "乙方统一社会信用代码/身份证号码校验未通过")
                    # factors_to_inform['乙方（卖方）统一社会信用代码/身份证号码审核提示'] = "请审核乙方（卖方）统一社会信用代码/身份证号码是否填写正确"
        else:
            factors_error["乙方（卖方）统一社会信用代码/身份证号码"] = "统一社会信用代码/身份证号码未填写完整"
            # factors_to_inform['乙方（卖方）统一社会信用代码/身份证号码审核提示'] = "请审核乙方（卖方）统一社会信用代码/身份证号码是否填写正确"
            hp.addRemarkInDoc(word, document, "乙方（卖方）", "乙方统一社会信用代码/身份证号码未填写")
    except:
        factors_miss.append("乙方（卖方）统一社会信用代码/身份证号码合同要素缺失")
        flag = 0

    # 联系方式
    # 共三个，还有一个在合同履行方式下
    try:
        match = '联系方式：(.*?)\n'
        factor = re.findall(match, text)[1].replace(" ", "")
        factors["乙方（卖方）联系方式"] = factor
        if factor != "":
            factor = factor.replace("（", "").replace("）", "").replace("-", "")
            # if len(factor)==10:
            # factors_ok.append("乙方（卖方）联系方式")
            # hp.addRemarkInDoc(word, document, "联系方式", "联系方式为10位号码，请人工检查")
            # else:
            if isTelPhoneNumber(factor) != "Error":
                factors_ok.append("乙方（卖方）联系方式")
            else:
                factors_error["乙方（卖方）联系方式"] = "联系方式填写有误"
                # factors_to_inform['乙方（卖方）联系方式审核提示'] = "请审核乙方（卖方）联系方式是否填写正确"
                hp.addRemarkInDoc(word, document, "乙方（卖方）", "乙方（卖方）联系方式错误")
        else:
            factors_error["乙方（卖方）联系方式"] = "乙方（卖方）联系方式未填写完整"
            # factors_to_inform['乙方（卖方）联系方式审核提示'] = "请审核乙方（卖方）联系方式是否填写正确"
            hp.addRemarkInDoc(word, document, "乙方（卖方）", "乙方（卖方）联系方式不完整")
    except:
        factors_miss.append("乙方（卖方）联系方式合同要素缺失")
        flag = 0

    # 邮箱
    try:
        match = '电子邮箱：(.*?)\n'
        factor = re.findall(match, text)[1].replace(" ", "")
        factors["乙方（卖方）电子邮箱"] = factor
        if factor != "":
            if checkEmail(factor):
                factors_ok.append("乙方（卖方）电子邮箱")
            else:
                factors_error["乙方（卖方）电子邮箱"] = "乙方（卖方）电子邮箱填写有误"
                hp.addRemarkInDoc(word, document, "乙方（卖方）", "乙方（卖方）电子邮箱错误")
                # factors_to_inform['乙方（卖方）电子邮箱'] = "乙方（卖方）电子邮箱填写有误"
        else:
            factors_error["乙方（卖方）电子邮箱"] = "乙方（卖方）电子邮箱未填写完整"
            hp.addRemarkInDoc(word, document, "乙方（卖方）", "乙方（卖方）电子邮箱不完整")
            # factors_to_inform['乙方（卖方）电子邮箱'] = "乙方（卖方）电子邮箱填写有误"
    except:
        factors_miss.append("乙方（卖方）电子邮箱合同要素缺失")
        flag = 0

    nl = '\n·'
    if flag == 0:
        factors_miss_block.append(
            f"\n风险名称：合同主体信息不完善\n风险偏向：(全局)不利于买卖双方\n风险提示：主体信息应包括：买/卖方姓名（名称）、法定代表人/负责人、住所地、统一社会信用代码/身份证号码、联系电话、电子邮箱\n缺失的要素：{nl}{nl.join(factors_miss)}")
        # addRemarkInDoc(word,document,"\n",f"\n风险名称：合同主体信息不完善\n风险偏向：(全局)不利于买卖双方\n"
        # f"风险提示：主体信息应包括：买/卖方姓名（名称）、法定代表人/负责人、住所地、统一社会信用代码/身份证号码、联系电话、电子邮箱\n缺失的要素：{nl}{nl.join(factors_miss)}")

    print(factors_miss)
    factors_miss = []
    print(factors_miss)

    flag = 1
    # -------------乙方检查完毕----------------#

    # 1.订购货品情况审查
    # 货品清单检查
    # 提示相关部门人工审查合同情况，有错误，或为空值，不另外提示；全对则提示相关部门人工审查
    # factors_to_inform['货品情况'] = "请审核订购货品情况（货品清单）是否与双方约定或招标文件一致"
    try:
        match = '货品名称：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["货品名称"] = factor
        if factor != '':
            factors_ok.append("货品名称")
            goods_flag_1 = True
            # factors_to_inform['货品名称审核提示'] = "请审核货品名称是否与招标文件一致"
        else:
            factors_error['货品名称'] = "货品名称未填写完整"
            hp.addRemarkInDoc(word, document, "货品名称", "货品名称未填写完整")
            goods_flag_1 = False
            # factors_to_inform['货品名称审核提示'] = "请审核货品名称是否与招标文件一致"
    except:
        factors_miss.append("货品名称合同要素缺失")
        flag = 0

    # str_miss = "；".join(factors_miss)
    # print(str_miss)
    # hp.addRemarkInDoc(word,document,"\n",str_miss)

    try:
        match = '货品型号：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["货品型号"] = factor
        if factor != '':
            factors_ok.append("货品型号")
            goods_flag_2 = True
            # factors_to_inform['货品型号审核提示'] = "请审核货品型号是否与招标文件一致"
        else:
            factors_error['货品型号'] = "货品型号未填写完整"
            hp.addRemarkInDoc(word, document, "货品型号", "货品型号未填写完整")
            goods_flag_2 = False
            # factors_to_inform['货品型号审核提示'] = "请审核货品型号是否与招标文件一致"
    except:
        factors_miss.append("货品型号合同要素缺失")
        flag = 0

    try:
        match = '货品材质：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["货品材质"] = factor
        if factor != '':
            factors_ok.append("货品材质")
            goods_flag_3 = True
            # factors_to_inform['货品材质审核提示'] = "请审核货品材质是否与招标文件一致"
        else:
            factors_error['货品材质'] = "货品材质未填写完整"
            hp.addRemarkInDoc(word, document, "货品材质", "货品材质未填写完整")
            goods_flag_3 = False
            # factors_to_inform['货品材质审核提示'] = "请审核货品材质是否与招标文件一致"
    except:
        factors_miss.append("货品材质合同要素缺失")
        flag = 0

    try:
        match = '规格参数：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["规格参数"] = factor
        if factor != '':
            factors_ok.append("规格参数")
            goods_flag_4 = True
            # factors_to_inform['规格参数审核提示'] = "请审核规格参数是否与招标文件一致"
        else:
            factors_error['规格参数'] = "规格参数未填写完整"
            hp.addRemarkInDoc(word, document, "规格参数", "规格参数未填写完整")
            goods_flag_4 = False
            # factors_to_inform['规格参数审核提示'] = "请审核规格参数是否与招标文件一致"
    except:
        factors_miss.append("规格参数合同要素缺失")
        flag = 0

    try:
        match = '货品数量（单位）：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["货品数量（单位）"] = factor
        if factor != '':
            factors_ok.append("货品数量（单位）")
            goods_flag_5 = True
            # factors_to_inform['货品数量审核提示'] = "请审核货品数量是否与招标文件一致"
        else:
            factors_error['货品数量（单位）'] = "货品数量（单位）未填写完整"
            hp.addRemarkInDoc(word, document, "货品数量", "货品数量未填写完整")
            goods_flag_5 = False
            # factors_to_inform['货品数量审核提示'] = "请审核货品数量是否与招标文件一致"
    except:
        factors_miss.append("货品数量合同要素缺失")
        flag = 0

    try:
        match = '货品单价（元）：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["货品单价（元）"] = factor
        if factor != '':
            factors_ok.append("货品单价（元）")
            goods_flag_6 = True
        else:
            factors_error['货品单价（元）'] = "货品单价（元）未填写完整"
            hp.addRemarkInDoc(word, document, "货品单价", "货品单价未填写完整")
            goods_flag_6 = False
            # factors_to_inform['货品单价审核提示'] = "请审核货品单价是否与招标文件一致"
    except:
        factors_miss.append("货品单价（元）合同要素缺失")
        flag = 0

    try:
        match = '其他：([\s\S]*?)二'
        factor = re.findall(match, text)[0].replace(" ", "").replace('\n', '')
        # print(factor)

        factors["其他"] = factor
        factor1 = re.split('（\d）', factor)

        while '' in factor1:
            factor1.remove("")

        # print(factor1)
        # print(len(factor1))
        if len(factor1) != 0:
            factors_ok.append("货品情况-其他")
            goods_flag_7 = True
        else:
            factors_error['其他'] = "货品情况-其他未填写完整"
            hp.addRemarkInDoc(word, document, "其他：", "货品情况-其他未填写完整")
            goods_flag_7 = False
    except:
        factors_miss.append("货品情况-其他合同要素缺失")
        flag = 0

    if goods_flag_1 and goods_flag_2 and goods_flag_3 and goods_flag_4 and goods_flag_5 and goods_flag_6 and goods_flag_7:
        factors_to_inform["订购货品情况"] = "请审核货品情况是否与双方约定或招标文件等一致"
        hp.addRemarkInDoc(word, document, "订购货品情况", "请审核货品情况是否与双方约定或招标文件等一致")

    # 其他是否需要检查？
    # code here or not

    # 2.货品价格总款审查
    # 货品价款总额
    try:
        match_total = '（价税合计）为人民币 *(\d*)'
        match_total_kanji = '大写： (.*?)元'  # 价格中文大写，用于检查大小写是否一致
        factor_total = re.findall(match_total, text)[0].replace(" ", "")

        factor_total_kanji = re.findall(match_total_kanji, text)[0]
        # 去掉圆整等汉字
        factor_total_kanji = factor_total_kanji.replace(" ", "").replace("整", "").replace("圆", "").replace("元", "")
        # print(factor_total_kanji)
        # print(type(factor_total_kanji))
        factors["货品总金额（价税合计）"] = factor_total
        # print(factor_total)
        if factor_total == '':
            # print("货品总金额根本没匹配到")
            factors_error['货品总金额（价税合计）'] = "货品总金额（价税合计）要素未填写"
            hp.addRemarkInDoc(word, document, "货品总金额", "货品总金额未填写完整")
            # factors_to_inform['货品总金额'] = "请审核货品总金额是否填写"
        else:
            container = hp.money_en_to_cn(float(factor_total)).replace("圆", "")
            # print(type(container))
            if container == factor_total_kanji:
                # 金额大小写能匹配
                factors_ok.append("货品总金额（价税合计）")
            else:
                factors_error['货品总金额（价税合计）'] = "货品总金额（价税合计）人民币大写与数值不一致"
                hp.addRemarkInDoc(word, document, "货品总金额", "货品总金额（价税合计）人民币大写与数值不一致")
                # factors_to_inform['货品总金额'] = "请审核货品总金额大小写是否一致"
    except:
        factors_miss.append("货品总金额（价税合计）合同要素缺失")
        flag = 0

    try:
        match = '合同价款（不含税）为人民币 *(\d*)'  # 此处括号为英文括号，标记一下
        match_kanji = '大写：(.*?) 元'  # 价格中文大写，用于检查大小写是否一致
        factor_total = re.findall(match, text)[0].replace(" ", "")
        factor_total_kanji = re.findall(match_kanji, text)[1].replace(" ", "").replace("整", "").replace("圆",
                                                                                                        "").replace("元",
                                                                                                                    "")
        # print(factor_total,factor_total_kanji)
        factors["合同价款(不含税)"] = factor_total
        if factor_total == '':
            factors_error['合同价款(不含税)'] = "合同价款(不含税)要素未填写"
            hp.addRemarkInDoc(word, document, "合同价款", "合同价款（不含税）未填写完整")
            # factors_to_inform['合同价款(不含税)'] = "请审核合同价款(不含税)是否填写"
        else:
            container = hp.money_en_to_cn(float(factor_total)).replace("圆", "")
            # print(container)
            if container == factor_total_kanji:
                # 金额大小写能匹配
                factors_ok.append("合同价款(不含税)")
            else:
                factors_error['合同价款(不含税)'] = "合同价款(不含税)人民币大写与数值不一致"
                hp.addRemarkInDoc(word, document, "合同价款", "合同价款（不含税）人民币大写与数值不一致")
                # factors_to_inform['合同价款(不含税)'] = "请审核合同价款(不含税)大小写是否一致"

    except:
        factors_miss.append("合同价款合同要素缺失")
        flag = 0

    try:
        match = '税金为人民币 *(\d*)'
        match_kanji = '大写：(.*?) 元'  # 价格中文大写，用于检查大小写是否一致
        factor_total = re.findall(match, text)[0].replace(" ", "")
        # print(factor_total)
        factor_total_kanji = re.findall(match_kanji, text)[2].replace(" ", "").replace("整", "").replace("圆",
                                                                                                        "").replace("元",
                                                                                                                    "")
        # print(factor_total,factor_total_kanji)
        factors["税金"] = factor_total
        if factor_total == '':
            factors_error['税金'] = "税金要素未填写"
            hp.addRemarkInDoc(word, document, "税金", "税金未填写完整")
            # factors_to_inform['税金'] = "请审核税金是否填写"
        else:
            container = hp.money_en_to_cn(float(factor_total)).replace("圆", "")
            # print(container)
            if container == factor_total_kanji:
                # 金额大小写能匹配
                factors_ok.append("税金")
            else:
                factors_error['税金'] = "税金人民币大写与数值不一致"
                hp.addRemarkInDoc(word, document, "税金", "税金大小写不一致")
                # factors_to_inform['税金'] = "请审核税金大小写是否一致"
    except:
        factors_miss.append("税金合同要素缺失")
        flag = 0

    try:
        match = '适用税率为 *(\d*)'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["税率"] = factor
        if factor != '':
            factors_ok.append("税率")
        else:
            factors_error['税率'] = "税率未未填写完整"
            hp.addRemarkInDoc(word, document, "适用税率", "未填写适用税率")
            # factors_to_inform['税率'] = "请审核税率是否填写"
    except:
        factors_miss.append("适用税率合同要素缺失")
        flag = 0

    if flag == 0:
        factors_miss_block.append(
            f"\n风险名称：货物（标的）信息不完善\n风险偏向：(全局)不利于买卖双方\n风险提示：货物（标的）信息应包括：货物（标的）名称、规格型号、计量单位、数量、产地、单价、税率、总价\n缺失的要素：{nl}{nl.join(factors_miss)}")
        # addRemarkInDoc(word, document, "\n", f"\n风险名称：货物（标的）信息不完善\n风险偏向：(全局)不利于买卖双方\n"
        #                                      f"风险提示：货物（标的）信息应包括：货物（标的）名称、规格型号、计量单位、数量、产地、单价、税率、总价\n缺失的要素：{nl}{nl.join(factors_miss)}")
    # print(factor)
    flag = 1
    factors_miss = []

    # 3.质量要求审查
    # 此处只审查了是否为空，未用正则审查国标
    # factors_to_inform['材质标准'] = "请审核材质标准、引用国家标准是否正确"
    try:
        match = '需符合 *(.*)所提供的材质标准'  # 此处冒号为英文冒号，特此标注
        factor = re.findall(match, text)[0].replace(" ", "")  # 找到主体名并去掉空格
        # print(factor)
        factors["材质标准"] = factor
        if factor != "":
            factors_ok.append("材质标准")
            matrial_flag_1 = True
            # factors_to_inform['材质标准'] = "请审核材质标准是否正确"
        else:
            factors_error["材质标准"] = "未填写材质标准"
            hp.addRemarkInDoc(word, document, "需符合", "材质标准未填写完整")
            matrial_flag_1 = False
            # factors_to_inform['材质标准'] = "请审核材质标准是否正确"
    except:
        factors_miss.append("材质标准合同要素缺失")

    try:
        match = '引用标准如下：([\s\S]*?)3.'
        factor = re.findall(match, text)[0].replace(" ", "").replace('\n', '')
        # print(factor)

        factors["其他"] = factor
        factor1 = re.split('（\d）', factor)

        while '' in factor1:
            factor1.remove("")

        # print(factor1)
        # print(len(factor1))
        if len(factor1) != 0:
            factors_ok.append("引用标准")
            matrial_flag_2 = True
        else:
            factors_error['引用标准'] = "引用标准未填写完整"
            hp.addRemarkInDoc(word, document, "引用标准如下", "引用标准未填写完整")
            matrial_flag_2 = True
    except:
        factors_miss.append("引用标准合同要素缺失")

    '''
    try:
        match = '引用标准如下：\n *(.*?)\n'  # 匹配多行？，此处仅匹配单行
        factor = re.findall(match, text)[0].replace(" ", "")  # 找到主体名并去掉空格
        # print(factor)
        factors["引用标准"] = factor
        if factor != "":
            factors_ok.append("引用标准")
            matrial_flag_2 = True
            #factors_to_inform['材质标准'] = "请审核引用标准是否正确"
        else:
            factors_error["引用标准"] = "未填写引用标准"
            hp.addRemarkInDoc(word, document, "引用标准", "引用标准未填写完整")
            matrial_flag_2 = False
            #factors_to_inform['引用标准'] = "请审核引用标准是否正确"
    except:
        factors_miss.append("引用标准合同要素缺失")
    '''

    if matrial_flag_1 and matrial_flag_2:
        factors_to_inform["质量要求"] = "请审核材质标准、引用国家标准是否正确"
        hp.addRemarkInDoc(word, document, "质量要求", "请审核材质标准、引用国家标准是否正确")

    # factors_to_inform['合同履行方式'] = "请审核合同履行方式细则是否与双方约定或招标文件一致"
    try:
        match = '交货地点：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["交货地点"] = factor
        if factor != '':
            factors_ok.append("交货地点")
            contractExecution_flag_1 = True

        else:
            factors_error["交货地点"] = "交货地点未填写"
            hp.addRemarkInDoc(word, document, "交货地点", "交货地点未填写完整")
            contractExecution_flag_1 = False
            # factors_to_inform['交货地点'] = "请审核交货地点是否填写"
    except:
        # factors_miss.append("交货地点合同要素缺失")
        factors_miss_block.append(f"\n风险名称：货物（标的）交付地点缺失\n风险偏向：(全局)不利于买卖双方\n风险提示：建议增加：买方指定货物交付地点为：【】")

    try:
        match = '交货日期：(.*?)年(.*?)月(.*?)日\n'
        factor = list(tuple(re.findall(match, text)[0]))
        # factors["交货日期"] = factor
        # factors_to_inform['交货日期'] = "请审核交货日期是否与招标文件一致"
        if len(factor) == 3:
            for i in range(3):
                factor[i] = factor[i].replace(' ', '')
            factors["交货日期"] = f"{factor[0]}-{factor[1]}-{factor[2]}"
            if hp.CheckDate(factor[0], factor[1], factor[2]):
                factors_ok.append("交货日期")
                contractExecution_flag_2 = True

            else:
                factors_error["交货日期"] = "交货日期填写不规范"
                hp.addRemarkInDoc(word, document, "交货日期", "交货日期填写不规范")
                contractExecution_flag_2 = False
                # factors_to_inform['交货日期'] = "请审核交货日期是否填写规范"
        else:
            factors["交货日期"] = factor
            factors_error["交货日期"] = "交货日期未填写完整"
            hp.addRemarkInDoc(word, document, "交货日期", "未填写交货日期")
            contractExecution_flag_2 = False
            # factors_to_inform['交货日期'] = "请审核交货日期是否填写"
    except:
        factors_miss.append("交货日期合同要素缺失")

    try:
        match = '交货方式： (.*?)\n'
        factor_deliveryType = re.findall(match, text)[0]
        factor_deliveryType = factor_deliveryType.replace(" ", "").replace("（自提|包送包卸货|包送包安装）（选项三选一）", "").replace(" ",
                                                                                                                  "")
        factors["交货方式"] = factor_deliveryType
        # factors_to_inform['交货方式'] = "请审核交货方式是否与招标文件一致"
        # factor_deliveryType=factor_deliveryType.replace(" ", "")
        # print(factor_deliveryType)
        # 判断是否三选一：
        if factor_deliveryType in ["自提", "包送包卸货", "包送包安装"]:
            factors["交货方式"] = f"{factor_deliveryType[0]}"
            factors_ok.append("交货方式")
            contractExecution_flag_3 = True
        else:
            factors_error["交货方式"] = "交货方式填写错误，请填写自提、包送包卸货或包送包安装"
            hp.addRemarkInDoc(word, document, "交货方式", "未选择交货方式或者非三项选择之一，请填写自提、包送包卸货或包送包安装")
            contractExecution_flag_3 = False
            # factors_to_inform['交货方式'] = "请审核交货方式是否填写"
    except:
        factors_miss.append("交货方式合同要素缺失")

        '''
        if factor_deliveryType != "":
            factors["交货方式"]=f"{factor_deliveryType[0]}"
            factors_ok.append("交货方式")
        else:
            factors_error["交货方式"]="交货方式未选择"
            hp.addRemarkInDoc(word, document, "交货方式", "要素填写错误：交货方式未选择")
        '''

    try:
        match = '交付标准：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["交付标准"] = factor
        # factors_to_inform['交付标准'] = "请审核交付标准是否与招标文件一致"
        if factor != '':
            factors_ok.append("交付标准")
            contractExecution_flag_4 = True
        else:
            factors_error["交付标准"] = "交付标准未填写"
            hp.addRemarkInDoc(word, document, "交付标准", "交付标准未填写完整")
            contractExecution_flag_4 = False
            # factors_to_inform['交付标准'] = "请审核交付标准是否填写"
    except:
        factors_miss.append("交付标准合同要素缺失")

    try:
        match = '验收标准：(.*?)\n'
        factor = re.findall(match, text)[0].replace(' ', '')
        # factors
    except:
        factors_miss_block.append(f"\n风险名称：货物（标的）验收标准缺失\n风险偏向：(全局)不利于买卖双方\n风险提示：建议增加货物（标的）验收标准")

    '''

    '''
    try:
        match = '联系人及联系方式：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["联系人及联系方式"] = factor
        # factors_to_inform['联系人及联系方式审核提示'] = "请审核联系人及联系方式是否与招标文件一致"
        # print(factor)
        if factor != "":
            factors_ok.append("联系人及联系方式")
            contractExecution_flag_5 = True

        else:
            # print("居然没找到？？")

            factors_error["联系人及联系方式"] = "联系人及联系方式未填写完整"
            # factors_to_inform['联系人及联系方式审核提示'] = "请审核联系人及联系方式是否填写正确"
            hp.addRemarkInDoc(word, document, "联系人及联系方式", "联系人及联系方式未填写完整")
            contractExecution_flag_5 = False
    except:
        # factors_miss.append("联系人及联系方式合同要素缺失")
        factors_miss_block.append(f"\n风险名称：收货联系人信息缺失\n风险偏向：(全局)不利于买卖双方\n"
                                  f"风险提示：建议增加：收货人姓名、收货人联系方式")

    # str_miss = "；".join(factors_miss)
    # print(str_miss)
    # hp.addRemarkInDoc(word, document, "\n", str_miss)

    try:
        match = '运输方式：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["运输方式"] = factor
        # factors_to_inform['运输方式'] = "请审核运输方式是否与招标文件一致"
        if factor != '':
            factors_ok.append("运输方式")
            contractExecution_flag_6 = True
        else:
            factors_error["运输方式"] = "运输方式未填写"
            hp.addRemarkInDoc(word, document, "运输方式", "运输方式未填写完整")
            contractExecution_flag_6 = False
            # factors_to_inform['运输方式'] = "请审核运输方式是否填写"
    except:
        # factors_miss.append("运输方式合同要素缺失")
        factors_miss_block.append(f"\n风险名称：货物（标的）运输信息不完善\n风险偏向：(全局)不利于买卖双方\n"
                                  f"风险提示：货物（标的）运输信息应包括：运输时间、装运地、运输方式、运费负担、运输保险费用承担、运输通知")

    if contractExecution_flag_1 and contractExecution_flag_2 and contractExecution_flag_3 and contractExecution_flag_4 and contractExecution_flag_5 and contractExecution_flag_6:
        factors_to_inform["合同履行方式"] = "请审核合同履行方式相关要素是否与双方约定或招标文件一致"
        hp.addRemarkInDoc(word, document, "合同履行方式", "请审核合同履行方式相关要素是否与双方约定或招标文件一致")

    # 验收期限
    try:
        match = '买方应在卖方(.*?)后'
        factor_receive = re.findall(match, text)[0].replace(" ", "")
        factors["验收期限-交货方式"] = factor_receive
        container = hp.DeliveryType(factor_deliveryType)
        if factor_receive == "":
            factors_error["验收期限-交货方式"] = "验收期限-交货方式未填写"
            hp.addRemarkInDoc(word, document, "买方应在卖方", "交货方式未填写完整")
            check_flag_1 = False
            # factors_to_inform['验收期限-交货方式'] = "请审核交货方式是否填写"
        else:
            if container == factor_receive:
                # factors["验收期限-交货方式"]=factor_receive
                factors_ok.append("验收期限-交货方式")
                check_flag_1 = True
            else:
                factors_error["验收期限-交货方式"] = "卖方交货方式与合同第四款第三项不一致"
                hp.addRemarkInDoc(word, document, "买方应在卖方", "交货方式与第四款第三项不一致")
                check_flag_1 = False
                # factors_to_inform['验收期限-交货方式'] = "请审核交货方式是否填写"
    except:
        # factors_miss.append("交货方式合同要素缺失")
        factors_miss_block.append(f"\n风险名称：货物（标的）交付方式缺失\n风险偏向：(全局)不利于买卖双方\n"
                                  f"风险提示：建议增加：双方约定本合同项下货物交付采用 【】（选填“一次性”或“分批次”）交付")

    # 此处尚未添加验收期限是否合理的判断，仅判断是否非空
    try:
        match = '最长不超过(.*?)个工作日'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["验收期限"] = factor
        factors_to_inform['验收期限'] = "请审核验收期限设置是否合理"
        if factor != '':
            factors_ok.append("验收期限")
            check_flag_2 = True
        else:
            factors_error["验收期限"] = "验收期限未填写"
            hp.addRemarkInDoc(word, document, "最长不超过", "验收期限填写错误")
            check_flag_2 = False
            # factors_to_inform['验收期限'] = "请审核验收期限是否填写"
    except:
        factors_miss.append("验收期限合同要素缺失")

    # 此处尚未添加签署验收文件期限是否合理的判断，仅判断是否非空
    try:
        match = '如买方在卖方安装完毕(.*?) 日'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["签署验收文件"] = factor
        factors_to_inform['签署验收文件'] = "请审核签署验收文件期限设置是否合理"
        if factor != '':
            factors_ok.append("签署验收文件")
            check_flag_3 = True
            # print("factor不是空的")
        else:
            # print("test")
            factors_error["签署验收文件"] = "签署验收文件期限未填写"
            hp.addRemarkInDoc(word, document, "卖方安装完毕", "签署验收文件期限填写错误")
            check_flag_3 = False
            # factors_to_inform['签署验收文件'] = "请审核安装完毕日期是否填写"
    except:
        factors_miss.append("签署验收文件期限合同要素缺失")

    if check_flag_1 and check_flag_2 and check_flag_3:
        factors_to_inform["验收期限"] = "请审核验收期限设置是否合理"
        hp.addRemarkInDoc(word, document, "验收期限", "请审核验收期限设置是否合理")

    # 保质期
    # 此处尚未添加货物保质期是否合理的判断，仅判断是否非空
    try:
        match = '货物保质期为(.*?)年'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["货物保质期"] = factor

        if factor != '':
            factors_ok.append("货物保质期")
            factors_to_inform['货物保质期'] = "请审核货物保质期设置是否符合相关规定、是否合理"
            hp.addRemarkInDoc(word, document, "货物保质期", "请审核保质期设置是否符合相关规定、是否合理")
            # keep_flag=True
        else:
            factors_error["货物保质期"] = "货物保质期未填写"
            hp.addRemarkInDoc(word, document, "货物保质期", "货物保质期未填写完整")
            # keep_flag = False
            # factors_to_inform['货物保质期'] = "请审核货物保质期是否填写"
    except:
        factors_miss.append("货物保质期合同要素缺失")

    # if keep_flag:
    # factors_to_inform["验收期限"] = "请审核验收期限设置是否合理"
    # hp.addRemarkInDoc(word, document, "验收期限", "请审核验收期限设置是否合理")

    # 5.付款方式及期限
    # 仅判断是否为空，需要招标文件进行进一步判断
    try:
        # match = '付款方式及期限：\n *(.*?)\n'  # 匹配单行
        match = '付款方式及期限：([\s\S]*?)六、'
        factor = re.findall(match, text)[0].replace(" ", "")  # 找到主体名并去掉空格
        print(factor)
        factors["付款方式及期限"] = factor
        factors_to_inform['付款方式及期限'] = "请审核付款方式及期限是否与双方约定或招标文件一致"
        if factor != "":
            factors_ok.append("付款方式及期限")
            factors_to_inform['付款方式及期限'] = "请审核付款方式及期限是否与双方约定或招标文件一致"
            hp.addRemarkInDoc(word, document, "付款方式及期限", "请审核付款方式与期限是否与双方约定或招标文件一致。")
            if "一次性" in factor:
                hp.addRemarkInDoc(word, document, '付款方式及期限', f"风险偏向：不利于买方\n风险提示：一次性支付全部货款将增加买方风险")
            elif "分期" in factor:
                hp.addRemarkInDoc(word, document, '付款方式及期限',
                                  f"风险偏向：不利于买卖双方\n风险提示：根据《中华人民共和国民法典》的规定，分期付款的买受人未支付到期价款的数额达到全部价款的五分之一，经催告后在合理期限内仍未支付到期价款的，出卖人可以请求买受人支付全部价款或者解除合同。出卖人解除合同的，可以向买受人请求支付该标的物的使用费。")
        else:
            factors_error["付款方式及期限"] = "付款方式及期限未填写"
            hp.addRemarkInDoc(word, document, "付款方式及期限", "付款方式及期限填写错误")
            # factors_to_inform['付款方式及期限'] = "请审核付款方式及期限是否填写"
    except:
        # factors_miss.append("付款方式及期限合同要素缺失")
        factors_miss_block.append(f"\n风险名称：货物（标的）交付方式缺失\n风险偏向：(全局)不利于买卖双方\n"
                                  f"风险提示：应将付款方式予以明确：例如分期付款、一次性付款、账期付款")

    # 6.甲乙双方账户信息
    # 乙方收款信息
    factors_miss = []
    flag = 1
    try:
        match = '纳税人类别：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["乙方纳税人类别"] = factor
        # factors_to_inform['乙方纳税人类别'] = "请审核乙方纳税人类别是否填写正确"
        # print(factor=="自然人")
        # print(factor == "个体商户")
        # print(factor == "法人")
        if factor != '':
            if factor == "自然人" or factor == "个体商户" or factor == "法人":
                # print("纳税人类别ok")
                factors_ok.append("乙方纳税人类别")
                account_flag_1 = True
            else:
                factors_error["乙方纳税人类别"] = "乙方纳税人类别填写错误"
                hp.addRemarkInDoc(word, document, "纳税人类别", "乙方纳税人类别填写错误，请填写自然人、个体商户或法人")
                account_flag_1 = False

        # factors_ok.append("乙方纳税人类别")
        else:
            factors_error["乙方纳税人类别"] = "乙方纳税人类别未填写完整"
            hp.addRemarkInDoc(word, document, "纳税人类别", "乙方纳税人类别未填写完整")
            account_flag_1 = False
            # factors_to_inform['乙方纳税人类别'] = "请审核乙方纳税人类别是否填写完整"
    except:
        factors_miss.append("乙方纳税人类别合同要素缺失")
        flag = 0

    try:
        match = '计税方式：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["乙方计税方式"] = factor
        # factors_to_inform['乙方计税方式'] = "请审核乙方计税方式是否填写正确"
        if factor != '':
            if factor == "一般计税方法" or factor == "简易计税方法" or factor == "扣税计税方法":
                factors_ok.append("乙方计税方式")
                account_flag_2 = True
            else:
                factors_error["乙方计税方式"] = "乙方计税方式填写错误"
                hp.addRemarkInDoc(word, document, "计税方式", "乙方计税方式填写错误，请填写一般计税法、简易计税法或扣税计税法")
                account_flag_2 = False

        # factors_ok.append("乙方纳税人类别")
        else:
            factors_error["乙方计税方式"] = "乙方计税方式未填写完整"
            hp.addRemarkInDoc(word, document, "计税方式", "乙方计税方式未填写完整，请填写一般计税法、简易计税法或扣税计税法")
            account_flag_2 = False
            # factors_to_inform['乙方计税方式'] = "请审核乙方计税方式是否填写"
    except:
        factors_miss.append("乙方计税方式合同要素缺失")
        flag = 0

    try:
        match = '纳税人识别号：(.*?)\n'
        social_match = '[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}'  # 纳税人识别号正则
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["乙方纳税人识别号"] = factor
        lenth = len(factor)
        # factors_to_inform['乙方纳税人识别号'] = "请审核乙方纳税人识别号是否填写正确"
        # print(lenth)
        # container = re.findall(social_match, factor)  # 首先检查社会信用代码
        # print(factor)
        # print(check_str('[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}',factor))
        # '['
        account_flag_3 = False
        if factor != "":
            if lenth == 15:
                if check_str('\d{6}[0-9A-Z]{9}', factor):
                    factors_ok.append("乙方纳税人识别号")
                    account_flag_3 = True
                else:
                    factors_error["乙方纳税人识别号"] = "乙方纳税人识别号填写错误"
                    hp.addRemarkInDoc(word, document, "纳税人识别号", "乙方纳税人识别号未通过校验")

            elif lenth == 17:
                if check_str('[0-9A-Z]{17}', factor):
                    factors_ok.append("乙方纳税人识别号")
                    account_flag_3 = True
                    # hp.addRemarkInDoc(word, document, "纳税人识别号", "乙方纳税人识别号未通过校验")
                    # factors_to_inform['乙方纳税人识别号'] = "请审核乙方纳税人识别号是否填写正确"
                else:
                    factors_error["乙方纳税人识别号"] = "乙方纳税人识别号填写错误"
                    hp.addRemarkInDoc(word, document, "纳税人识别号", "乙方纳税人识别号未通过校验")
                    # factors_to_inform['乙方纳税人识别号'] = "请审核乙方纳税人识别号是否填写正确"
            elif lenth == 18:
                if check_str('[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}', factor):
                    factors_ok.append("乙方纳税人识别号")
                    account_flag_3 = True
                else:
                    factors_error["乙方纳税人识别号"] = "乙方纳税人识别号填写错误"
                    hp.addRemarkInDoc(word, document, "纳税人识别号", "乙方纳税人识别号未通过校验")
                    # factors_to_inform['乙方纳税人识别号'] = "请审核乙方纳税人识别号是否填写正确"
            elif lenth == 20:
                if check_str('[0-9A-Z]{20}', factor):
                    factors_ok.append("乙方纳税人识别号")
                    account_flag_3 = True
                    # hp.addRemarkInDoc(word, document, "纳税人识别号", "乙方纳税人识别号位20位，请人工核对")
                    # factors_to_inform['乙方纳税人识别号'] = "请审核乙方纳税人识别号是否填写正确"
                else:
                    factors_error["乙方纳税人识别号"] = "乙方纳税人识别号填写错误"
                    hp.addRemarkInDoc(word, document, "纳税人识别号", "乙方纳税人识别号未通过校验")
                    # factors_to_inform['乙方纳税人识别号'] = "请审核乙方纳税人识别号是否填写正确"
            # if lenth==15 or 17 or 18 or 20:
            # if check_str('/^[A-Za-z0-9]+$/',factor):
            # factors_ok.append("乙方纳税人识别号")
            else:
                factors_error["乙方纳税人识别号"] = "乙方纳税人识别号填写错误"
                hp.addRemarkInDoc(word, document, "纳税人识别号", "乙方纳税人识别号未填写完整")
                # factors_to_inform['乙方纳税人识别号'] = "请审核乙方纳税人识别号是否填写正确"
        else:
            factors_error["乙方纳税人识别号"] = "乙方纳税人识别号未填写完整"
            hp.addRemarkInDoc(word, document, "纳税人识别号", "乙方纳税人识别号填写错误")
            # factors_to_inform['乙方纳税人识别号'] = "请审核乙方纳税人识别号是否填写正确"
    except:
        factors_miss.append("乙方纳税人识别号合同要素缺失")
        flag = 0

    try:
        match = '开户行：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")  # 找到主体名并去掉空格
        # print(factor)
        factors["乙方开户行"] = factor
        # factors_to_inform['乙方开户行'] = "请审核乙方开户行是否填写正确"
        if factor != "":
            factors_ok.append("乙方开户行")
            account_flag_4 = True
        else:
            factors_error["乙方开户行"] = "乙方开户行未填写"
            hp.addRemarkInDoc(word, document, "开户行", "未填写乙方开户行")
            account_flag_4 = False
            # factors_to_inform['乙方开户行'] = "请审核乙方开户行是否填写正确"
    except:
        factors_miss.append("乙方开户行合同要素缺失")
        flag = 0

    try:
        match = '开户名：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")  # 找到主体名并去掉空格
        # print(factor)
        factors["乙方开户名"] = factor
        # factors_to_inform['乙方开户名'] = "请审核乙方开户名是否填写正确"
        if factor != "":
            factors_ok.append("乙方开户名")
            account_flag_5 = True
        else:
            factors_error["乙方开户名"] = "乙方开户名未填写"
            hp.addRemarkInDoc(word, document, "开户名", "未填写乙方开户名")
            account_flag_5 = False
            # factors_to_inform['乙方开户名'] = "请审核乙方开户名是否填写正确"
    except:
        factors_miss.append("乙方开户名合同要素缺失")
        flag = 0

    try:
        match = '账号：(.*?)\n'
        match_bank_account = '^([1-9]{1})(\d{11}|\d{14}|\d{18}|\d{19})$'
        factor = re.findall(match, text)[0].replace(" ", "")  # 找到主体名并去掉空格
        # print(factor)
        # print(check_str(match_bank_account,factor))
        factors["乙方开户账号"] = factor
        # factors_to_inform['乙方开户账号'] = "请审核乙方开户账号是否填写"
        # container = re.findall(match_bank_account, factor)  # 检查银行账号格式是否正确，此处正则存疑
        # print(f"aaa+{container}+aaa")
        # print(type(container))
        # print(container!="")
        account_flag_6 = False
        if factor == "":
            factors_error["乙方开户账号"] = "乙方开户账号未填写"
            hp.addRemarkInDoc(word, document, "账号", "未填写乙方开户账号")
            # factors_to_inform['乙方开户账号'] = "请审核乙方开户账号是否填写"
        else:
            # if container != "":
            if check_str(match_bank_account, factor):
                factors_ok.append("乙方开户账号")
                account_flag_6 = True
                # print("判断错误")
            else:
                factors_error["乙方开户账号"] = "乙方开户账号填写错误"
                hp.addRemarkInDoc(word, document, "账号", "乙方开户账号填写错误")
                # factors_to_inform['乙方开户账号'] = "请审核乙方开户账号是否填写正确"
    except:
        factors_miss.append("乙方开户账号合同要素缺失")
        flag = 0

    if account_flag_1 and account_flag_2 and account_flag_3 and account_flag_4 and account_flag_5 and account_flag_6:
        factors_to_inform["乙方收款信息"] = "请审核乙方收款信息是否正确"
        hp.addRemarkInDoc(word, document, "乙方收款信息", "请审核乙方收款信息是否正确")

    # 甲方开票信息
    try:
        match = '纳税人类别：(.*?)\n'
        factor = re.findall(match, text)[1].replace(" ", "")
        factors["甲方纳税人类别"] = factor
        # factors_to_inform['甲方纳税人类别'] = "请审核甲方纳税人类别是否填写正确"
        account1_flag_1 = False
        if factor != '':
            if factor == "自然人" or factor == "个体商户" or factor == "法人":
                factors_ok.append("甲方纳税人类别")
                account1_flag_1 = True
            else:
                factors_error["甲方纳税人类别"] = "甲方纳税人类别填写错误"
                hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方纳税人类别填写错误，请填写自然人、个体商户或法人")
                # factors_to_inform['甲方纳税人类别'] = "请审核甲方纳税人类别是否填写正确"
        else:
            factors_error["甲方纳税人类别"] = "甲方纳税人类别未填写完整"
            hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方纳税人类别未填写完整，请填写自然人、个体商户或法")
            # factors_to_inform['甲方纳税人类别'] = "请审核甲方纳税人类别是否填写"
    except:
        factors_miss.append("甲方纳税人类别合同要素缺失")
        flag = 0

    try:
        match = '计税方式：(.*?)\n'
        factor = re.findall(match, text)[1].replace(" ", "")
        # print(factor)
        factors["甲方计税方式"] = factor
        # factors_to_inform['甲方计税方式'] = "请审核甲计税方式是否填写正确"
        account1_flag_2 = False
        if factor != '':
            if factor == "一般计税方法" or factor == "简易计税方法" or factor == "扣税计税方法":
                factors_ok.append("甲方计税方式")
                account1_flag_2 = True
            else:
                factors_error["甲方计税方式"] = "甲方计税方式填写错误"
                hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方计税方式填写错误，请填写一般计税法、简易计税法或扣税计税法")
                # factors_to_inform['甲方计税方式'] = "请审核甲计税方式是否填写正确"
        # factors_ok.append("乙方纳税人类别")
        else:
            factors_error["甲方计税方式"] = "甲方计税方式未填写完整"
            hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方计税方式未填写完整，请填写一般计税法、简易计税法或扣税计税法")
            # factors_to_inform['甲方计税方式'] = "请审核甲方计税方式是否填写"
    except:
        factors_miss.append("甲方计税方式合同要素缺失")
        flag = 0

    try:
        match = '纳税人识别号：(.*?)\n'
        # social_match = '[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}'  # 纳税人识别号正则
        factor = re.findall(match, text)[1].replace(" ", "")
        factors["甲方纳税人识别号"] = factor
        lenth = len(factor)
        # factors_to_inform['甲方纳税人识别号'] = "请审核甲方纳税人识别号是否填写正确"
        account1_flag_3 = False
        # print(factor)
        # container = re.findall(social_match, factor)  # 首先检查社会信用代码
        # print(type(factor))
        # print(lenth)

        # print(check_str('[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}',factor))
        if factor != "":
            if lenth == 15:
                if check_str('\d{6}[0-9A-Z]{9}', factor):
                    factors_ok.append("甲方纳税人识别号")
                    account1_flag_3 = True
                else:
                    factors_error["甲方纳税人识别号"] = "甲方纳税人识别号填写错误"
                    hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方纳税人识别号未通过校验")
                    # factors_to_inform['甲方纳税人识别号'] = "请审核甲方纳税人识别号是否填写正确"
            elif lenth == 17:
                if check_str('[0-9A-Z]{17}', factor):
                    factors_ok.append("甲方纳税人识别号")
                    account1_flag_3 = True

                else:
                    factors_error["甲方纳税人识别号"] = "甲方纳税人识别号填写错误"
                    hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方纳税人识别号未通过校验")
                    # factors_to_inform['甲方纳税人识别号'] = "请审核甲方纳税人识别号是否填写正确"
            elif lenth == 18:
                if check_str('[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}', factor):
                    factors_ok.append("甲方纳税人识别号")
                    account1_flag_3 = True
                else:
                    factors_error["甲方纳税人识别号"] = "甲方纳税人识别号填写错误"
                    hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方纳税人识别号未通过校验")
                    # factors_to_inform['甲方纳税人识别号'] = "请审核甲方纳税人识别号是否填写正确"
            elif lenth == 20:
                if check_str('[0-9A-Z]{20}', factor):
                    factors_ok.append("甲方纳税人识别号")
                    account1_flag_3 = True
                else:
                    factors_error["甲方纳税人识别号"] = "甲方纳税人识别号填写错误"
                    hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方纳税人识别号未通过校验")
                    # factors_to_inform['甲方纳税人识别号'] = "请审核甲方纳税人识别号是否填写正确"
            else:
                factors_error["甲方纳税人识别号"] = "甲方纳税人识别号填写错误"
                hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方纳税人识别号填写错误")
                # factors_to_inform['甲方纳税人识别号'] = "请审核甲方纳税人识别号是否填写正确"
        else:
            factors_error["甲方纳税人识别号"] = "甲方方纳税人识别号未填写完整"
            hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方纳税人识别号未填写完整")
            # factors_to_inform['甲方纳税人识别号'] = "请审核甲方纳税人识别号是否填写正确"
    except:
        factors_miss.append("甲方纳税人识别号合同要素缺失")
        flag = 0

    try:
        match = '开户行：(.*?)\n'
        factor = re.findall(match, text)[1].replace(" ", "")  # 找到主体名并去掉空格
        # print(factor)
        factors["甲方开户行"] = factor
        # factors_to_inform['甲方开户行'] = "请审核甲方开户行是否填写正确"
        account1_flag_4 = False
        if factor != "":
            factors_ok.append("甲方开户行")
            account1_flag_4 = True
        else:
            factors_error["甲方开户行"] = "甲方开户行未填写"
            hp.addRemarkInDoc(word, document, "甲方开票信息", "未填写甲方开户行")
            # factors_to_inform['甲方开户行'] = "请审核甲方开户行是否填写正确"
    except:
        factors_miss.append("甲方开户行合同要素缺失")

    try:
        match = '开户名：(.*?)\n'
        factor = re.findall(match, text)[1].replace(" ", "")  # 找到主体名并去掉空格
        # print(factor)
        factors["甲方开户名"] = factor
        # factors_to_inform['甲方开户名'] = "请审核甲方开户名是否填写正确"
        account1_flag_5 = False
        if factor != "":
            factors_ok.append("甲方开户名")
            account1_flag_5 = True
        else:
            factors_error["甲方开户名"] = "甲方开户名未填写"
            hp.addRemarkInDoc(word, document, "甲方开票信息", "未填写甲方开户名")
            # factors_to_inform['甲方开户名'] = "请审核甲方开户名是否填写正确"
    except:
        factors_miss.append("甲方开户名合同要素缺失")
        flag = 0

    try:
        match = '账号：(.*?)\n'
        match_bank_account = '^([1-9]{1})(\d{11}|\d{14}|\d{18}|\d{19})$'
        factor = re.findall(match, text)[1].replace(" ", "")  # 找到主体名并去掉空格
        # print(factor)
        factors["甲方开户账号"] = factor
        # factors_to_inform['甲方开户账号'] = "请审核甲方开户账号是否填写正确"
        account1_flag_6 = False
        # print(factor)
        # print(check_str(match_bank_account, factor))
        # container = re.findall(match_bank_account, factor)  # 检查银行账号格式是否正确，此处正则存疑
        if factor == "":
            factors_error["甲方开户账号"] = "甲方开户账号未填写"
            hp.addRemarkInDoc(word, document, "甲方开票信息", "未填写甲方开户账号")
            # factors_to_inform['甲方开户账号'] = "请审核甲方开户账号是否填写正确"
        else:
            if check_str(match_bank_account, factor):
                factors_ok.append("甲方开户账号")
                account1_flag_6 = True
            else:
                factors_error["甲方开户账号"] = "甲方开户账号填写错误"
                hp.addRemarkInDoc(word, document, "甲方开票信息", "甲方开户账号填写错误")
                # factors_to_inform['甲方开户账号'] = "请审核甲方开户账号是否填写正确"
    except:
        factors_miss.append("甲方开户账号合同要素缺失")
        flag = 0

    if account1_flag_1 and account1_flag_2 and account1_flag_3 and account1_flag_4 and account1_flag_5 and account1_flag_6:
        factors_to_inform["甲方开票信息"] = "请审核甲方开票信息是否正确"
        hp.addRemarkInDoc(word, document, "甲方开票信息", "请审核甲方开票信息是否正确")

    if flag == 0:
        factors_miss_block.append(
            f"\n风险名称：开票信息不完善\n风险偏向：(全局)不利于买卖双方\n风险提示：开票信息应包括：发票类型、公司名称、纳税识别号、地址、电话、开户银行、银行账户\n缺失的要素：{nl}{nl.join(factors_miss)}")

    # 7.售后服务
    # todo：需要招标文件以判断时限是否相符
    try:
        match = '买方维修要求后(.*?)小时'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["售后服务-回应时限"] = factor
        feedback_flag_1 = False
        # factors_to_inform['售后服务-回应时限'] = "请审核售后服务-回应时限设置是否与招标文件一致"
        if factor != "":
            factors_ok.append("售后服务-回应时限")
            feedback_flag_1 = True
        else:
            factors_error["售后服务-回应时限"] = "售后服务-回应时限未填写"
            hp.addRemarkInDoc(word, document, "维修要求后", "售后服务-回应时限未约定")

    except:
        factors_miss.append("售后服务-回应时限合同要素缺失")

    try:
        match = '回应，(.*?)小时'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["售后服务-到达时限"] = factor
        # factors_to_inform['售后服务-到达时限'] = "请审核售后服务-到达时限设置是否与招标文件一致"
        feedback_flag_2 = False
        if factor != "":
            factors_ok.append("售后服务-到达时限")
            feedback_flag_2 = True
        else:
            factors_error["售后服务-到达时限"] = "售后服务-到达时限未填写"
            hp.addRemarkInDoc(word, document, "回应，", "售后服务-到达时限未约定")
            # factors_to_inform['售后服务-到达时限'] = "请审核售后服务-到达时限是否填写正确"

    except:
        factors_miss.append("售后服务-到达时限合同要素缺失")

    try:
        match = '现场，(.*?)小时'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["售后服务-修复时限"] = factor
        # factors_to_inform['售后服务-修复时限'] = "请审核售后服务-修复时限设置是否与招标文件一致"
        feedback_flag_3 = False
        if factor != "":
            factors_ok.append("售后服务-修复时限")
            feedback_flag_3 = True
        else:
            factors_error["售后服务-修复时限"] = "售后服务-修复时限未填写"
            hp.addRemarkInDoc(word, document, "现场，", "售后服务-修复时限未约定")
            # factors_to_inform['售后服务-修复时限'] = "请审核售后服务-修复时限是否填写正确"
    except:
        factors_miss.append("售后服务-修复时限合同要素缺失")

    try:
        match = '承诺如果在(.*?)小时'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["售后服务-承诺修复"] = factor
        # factors_to_inform['售后服务-承诺修复'] = "请审核售后服务-陈诺修复时限设置是否与招标文件一致"
        feedback_flag_4 = False
        if factor != "":
            factors_ok.append("售后服务-承诺修复")
            feedback_flag_4 = True
        else:
            factors_error["售后服务-承诺修复"] = "售后服务-承诺修复时限未填写"
            hp.addRemarkInDoc(word, document, "卖方承诺如果在", "售后服务-承诺修复时限未约定")
            # factors_to_inform['售后服务-承诺修复'] = "请审核售后服务-陈诺修复时限是否填写正确"
    except:
        factors_miss.append("售后服务-承诺修复时限合同要素缺失")

    try:
        match = '修复，将在(.*?)小时'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["售后服务-延长修复"] = factor
        feedback_flag_5 = False
        # factors_to_inform['售后服务-延长修复时限'] = "请审核售后服务-延长修复时限设置是否与招标文件一致"
        if factor != "":
            factors_ok.append("售后服务-延长修复")
            feedback_flag_5 = True
        else:
            factors_error["售后服务-延长修复"] = "售后服务-延长修复时限未填写"
            hp.addRemarkInDoc(word, document, "将在", "售后服务-延长修复时限未约定")
            # factors_to_inform['售后服务-延长修复时限'] = "请审核售后服务-延长修复时限是否填写正确"
    except:
        factors_miss.append("售后服务-延长修复时限合同要素缺失")

    if feedback_flag_1 and feedback_flag_2 and feedback_flag_3 and feedback_flag_4 and feedback_flag_5:
        factors_to_inform["售后服务"] = "请审核时间设置是否与招投标文件一致"
        hp.addRemarkInDoc(word, document, "售后服务：", "请审核时间设置是否与招标文件一致")

    try:
        match = '违约责任：(.*?)\n'
        factor = re.findall(match, text).replace(" ", "")
    except:
        factors_miss_block.append(
            f"\n风险名称：违约责任条款约定不完善\n风险偏向：(全局)不利于买卖双方\n风险提示：违约责任条款应为：\n1、卖方逾期交付的（包括但不限于未能按时交付本合同约定的文件及资料、未在本合同约定的时间内提供安装及调试指导技术服务等），卖方应按逾期交付、逾期安装部分的货物对应合同款项【1】‰/天的比例向买方支付违约金）。超过交付日期【30】日，卖方仍然不能履行交货义务的，买方有权选择单方面解除本合同，卖方应退还全部已收款项，并按合同总金额的20％向买方支付违约金。\n2、货物在安装完毕后，双方一起进行调试运行验收。因货物质量问题或安装等原因出现差错，由卖方完全负责。货物质量和安装工程质量应达到国家或专业的质量检验评定标准的合格条件，达不到约定条件的部分，导致验收未能通过的，买方有权要求卖方限期整改，卖方应按买方要求整改，整改视为延误，每逾期一日，向买方支付合同货物总价款【1】‰作为违约金。整改期届满仍然未能通过验收或者未能按期整改完毕的，买方有权解除本合同，卖方应退还全部已收款项，并按合同总金额的20％向买方支付违约金\n3、除本合同另有约定外，任何一方均无权擅自单方解除合同。若任何一方有违法单方解除合同行为的，无论解除行为是否生效，均须向另一方支付违约金。违约金为本合同总金额的20%。\n4、卖方保证所提供的货物或其任何一个组成部分均不会侵犯任何第三方的专利权、商标权、著作权、商业秘密等合法权利；如出现侵权情形，买方有权解除本合同，卖方应退还全部已收款项，并按合同总金额的20％向买方支付违约金。\n5、双方均认可，本合同约定的违约金，包括补偿性违约金及惩罚性违约金，相对于违约可能给对方带来的损失，不属于畸高。违约金不足以赔偿守约方损失的，违约方还应据实赔偿损失，损失包括预期利益损失及其他直接、间接损失。\n6、双方应按本合同约定，履行各自的义务。如因一方不按本合同约定履行义务或延迟履行义务，造成另一方损失的（包括但不限于对违约后果采取补救措施而花费的费用，因违约行为导致第三方提出的索赔要求，以及因此而发生的调查费用、差旅费用、律师费用、公证费、鉴定费、及其他为完成举证责任而发生的费用等），违约方应承担赔偿责任。")

    # 8.违约责任
    try:
        match = '合同总价款的(.*?)%'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["卖方违约金"] = factor
        # factors_to_inform['卖方违约金'] = "请审核卖方违约金是否填写正确，比例是否<30%"
        if factor == "":
            factors_error["卖方违约金"] = "卖方违约金未填写"
            hp.addRemarkInDoc(word, document, "合同总价款", "未约定卖方违约金")
            # factors_to_inform['卖方违约金'] = "请审核卖方违约金是否填写"
        else:
            if int(factor) < 30:
                factors_ok.append("卖方违约金")
            else:
                factors_error["卖方违约金"] = "卖方违约金比例超出"
                hp.addRemarkInDoc(word, document, "合同总价款", "卖方违约金比例超出规定，应<30%")
                # factors_to_inform['卖方违约金'] = "请审核卖方违约金是否填写正确，比例是否<30%"
    except:
        factors_miss.append("卖方违约金合同要素缺失")

    try:
        match = '逾期超过(.*?)日'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["卖方逾期"] = factor
        if factor != "":
            factors_ok.append("卖方逾期")
        else:
            factors_error["卖方逾期"] = "卖方逾期未填写"
            hp.addRemarkInDoc(word, document, "逾期超过", "卖方逾期内容缺失")
            # factors_to_inform['卖方逾期'] = "请审核卖方逾期时限是否填写"
    except:
        factors_miss.append("卖方逾期合同要素缺失")

    try:
        match = '合同金额的(.*?)%'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["卖方违约金-逾期"] = factor
        # factors_to_inform['卖方违约金-逾期'] = "请审核逾期导致卖方违约金是否填写，比例是否<30%"
        if factor == "":
            factors_error["卖方违约金-逾期"] = "卖方违约金-逾期未填写"
            hp.addRemarkInDoc(word, document, "合同金额的", "未约定逾期导致的卖方违约金")
            # factors_to_inform['卖方违约金-逾期'] = "请审核卖方违约金是否填写"
        else:
            if int(factor) < 30:
                factors_ok.append("卖方违约金-逾期")
            else:
                factors_error["卖方违约金-逾期"] = "卖方违约金-逾期比例超出"
                hp.addRemarkInDoc(word, document, "合同金额的", "卖方违约金-逾期比例超出，应<30%")
                # factors_to_inform['卖方违约金-逾期'] = "请审核卖方违约金比例是否<30%"
    except:
        factors_miss.append("卖方违约金-逾期合同要素缺失")

    try:
        match = '本合同总金额(.*?)%'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["卖方违约金-单方面取消"] = factor
        # factors_to_inform['卖方违约金-单方面取消'] = "请审核卖方违约金（单方面取消）是否填写，比例是否<30%"
        if factor == "":
            factors_error["卖方违约金-单方面取消"] = "卖方违约金-单方面取消未填写"
            hp.addRemarkInDoc(word, document, "合同总金额", "未填写卖方违约金（单方面取消）")
            # factors_to_inform['卖方违约金-单方面取消'] = "请审核卖方违约金是否填写"
        else:
            if int(factor) < 30:
                factors_ok.append("卖方违约金-单方面取消")
            else:
                factors_error["卖方违约金-单方面取消"] = "卖方违约金-单方面取消比例超出"
                hp.addRemarkInDoc(word, document, "合同总金额", "卖方违约金-单方面取消比例超出规定，应<30%")
                # factors_to_inform['卖方违约金-单方面取消'] = "请审核卖方违约金比例是否<30%"
    except:
        factors_miss.append("卖方违约金-单方面取消合同要素缺失")

    if "通知送达" in text:
        pass
    else:
        factors_miss_block.append(f"\n风险名称：通知送达条款缺失\n风险偏向：(全局)不利于买卖双方\n"
                                  f"风险提示：还应包含以下送达条款：合同中应当约定买卖双方联系人、电话、地址、邮箱等信息。")

    if "保密条款" in text:
        pass
    else:
        factors_miss_block.append(f"\n风险名称：保密条款缺失\n风险偏向：(全局)不利于买卖双方\n"
                                  f"风险提示：建议增加以下条款：合同一方为履行本合同向对方提供的所有商业秘密、技术信息、产品知识产权等以及由披露方提供的第三方数据或信息，接收方未经披露方书面许可不得做其他用途以且不得披露或转让给任何第三方。")

    if "知识产权" in text:
        pass
    else:
        factors_miss_block.append(f"\n风险名称：知识产权条款缺失\n风险偏向：(全局)不利于买卖双方\n"
                                  f"风险提示：建议增加以下条款：所有交付成果如专为买方定制的成果，所有交付成果的产权、知识产权归属于买方。")

    if "不可抗力" in text:
        pass
    else:
        factors_miss_block.append(f"\n风险名称：不可抗力条款缺失\n风险偏向：(全局)不利于买卖双方\n"
                                  f"风险提示：建议增加以下条款：如果因合同一方无法合理控制的事由导致该合同方无法履行或迟延履行本合同项下的各项义务时，该合同方无须承担责任。")

    # todo：需要招标合同进行对比判断
    try:
        match = '规定的，应在(.*?)日'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["违约责任-免费更换期限"] = factor
        # factors_to_inform['违约责任-免费更换期限'] = "请审核更换日期是否与招标文件一致"
        if factor != "":
            factors_ok.append("违约责任-免费更换期限")
        else:
            factors_error["违约责任-免费更换期限"] = "违约责任-免费更换期限未填写"
            hp.addRemarkInDoc(word, document, "，应在", "违约责任-免费更换期限未约定")
            # factors_to_inform['违约责任-免费更换期限'] = "请审核期限是否填写"
    except:
        factors_miss.append("违约责任-免费更换期限合同要素缺失")

    try:
        match = '约定。超过(.*?)日'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["违约责任-免费更换超时"] = factor
        # factors_to_inform['违约责任-免费更换超时'] = "请审核期限是否与招标文件一致"
        if factor != "":
            factors_ok.append("违约责任-免费更换超时")
        else:
            factors_error["违约责任-免费更换超时"] = "违约责任-免费更换超时未填写"
            hp.addRemarkInDoc(word, document, "。超过", "违约责任-免费更换超时未约定")
            # factors_to_inform['违约责任-免费更换超时'] = "请审核期限是否填写"
    except:
        factors_miss.append("违约责任-免费更换超时合同要素缺失")

    try:
        match = '卖方向买方支付合同金额(.*?)%'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["卖方违约金-质量不符"] = factor
        # factors_to_inform['卖方违约金-质量不符'] = "请审核质量不符违约金比例是否填写，比例是否<30%"
        if factor == "":
            factors_error["卖方违约金-质量不符"] = "卖方违约金-质量不符未填写"
            hp.addRemarkInDoc(word, document, "卖方向买方支付合同金额", "未约定由于质量不符发生的卖方违约金")
            # factors_to_inform['卖方违约金-质量不符'] = "请审核金额比例是否填写"

        else:
            if int(factor) < 30:
                factors_ok.append("卖方违约金-质量不符")
            else:
                factors_error["卖方违约金-质量不符"] = "卖方违约金-质量不符比例超出"
                hp.addRemarkInDoc(word, document, "卖方向买方支付合同金额", "质量不符发生的卖方违约金比例超出规定，应<30%")

    except:
        factors_miss.append("卖方违约金-质量不符合同要素缺失")

    # 9.解决纠纷的方式
    try:
        match = '双方均可向(.*?)人'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["解决纠纷-甲方住所地"] = factor
        if factor == "":
            factors_error["解决纠纷-甲方住所地"] = "解决纠纷-甲方住所地未填写"
            hp.addRemarkInDoc(word, document, "均可向", "甲方住所地缺失")
            # factors_to_inform['解决纠纷-甲方住所地'] = "请审核甲方住所地是否填写"
        # 直接填甲方住所地的情况
        elif factor == "甲方住所地":
            factors_ok.append("解决纠纷-甲方住所地")
        else:
            # 拆分甲方住所地，以实现法院判定，factor_location拆分为省市区县
            # print(factor)
            # print(factor_location)
            factor_location_1 = factor_location.split("省")
            factor_location_2 = factor_location.split("市")
            factor_location_3 = factor_location.split("区")
            factor_location_4 = factor_location.split("县")
            # print(factor_location_1)
            # print(factor_location_2)
            # print(factor_location_3)
            # print(factor_location_4)
            flag1 = "省" in factor_location
            flag2 = "市" in factor_location
            flag3 = "区" in factor_location
            flag4 = "县" in factor_location
            # print(flag1)
            # print(flag2)
            # print(flag3)
            # print(flag4)
            # print((factor == factor_location_1[0] + "省" or factor_location_1[0] + "省高级" or factor_location_1[0] + "省高级人民" \
            # or factor_location_2[0] + "市" or factor_location_2[0] + "市中级" or factor_location_2[0] + "市中级人民" or factor_location_2[0] + "市高级" or factor_location_2[0] + "市高级人民" \
            # or factor_location_3[0] + "区" or factor_location_2[0] + "区人民"))

            # 如果地址填到县级，则判断到县级法院及上级法院
            if flag4:
                # if factor==factor_location_1[0]+"省"  or factor==factor_location_2[0]+"市" or factor==factor_location_3[0]+"区" or factor==factor_location_4[0]+"县":
                if factor == factor_location_1[0] + "省" or factor == factor_location_1[0] + "省高级" or factor == \
                        factor_location_1[0] + "省高级人民" \
                        or factor == factor_location_2[0] + "市" or factor == factor_location_2[0] + "市中级" or factor == \
                        factor_location_2[0] + "市中级人民" or factor == factor_location_2[0] + "市高级" or factor == \
                        factor_location_2[0] + "市高级人民" \
                        or factor == factor_location_3[0] + "区" or factor == factor_location_2[0] + "区人民" or \
                        factor == factor_location_4[0] + "县" or factor == factor_location_4[0] + "县人民":
                    factors_ok.append("解决纠纷-甲方住所地")
                else:
                    factors_error["解决纠纷-甲方住所地"] = "甲方住所地与合同页中描述不符"
                    hp.addRemarkInDoc(word, document, "均可向", "甲方住所地法院与合同页中描述不符")
                    # factors_to_inform['解决纠纷-甲方住所地'] = "请审核甲方住所地是否填写"

            elif flag3:
                # if factor==factor_location_1[0]+"省" or factor==factor_location_2[0]+"市" or factor==factor_location_3[0]+"区":
                if factor == factor_location_1[0] + "省" or factor == factor_location_1[0] + "省高级" or factor == \
                        factor_location_1[0] + "省高级人民" \
                        or factor == factor_location_2[0] + "市" or factor == factor_location_2[0] + "市中级" or factor == \
                        factor_location_2[0] + "市中级人民" or factor == factor_location_2[0] + "市高级" or factor == \
                        factor_location_2[0] + "市高级人民" \
                        or factor == factor_location_3[0] + "区" or factor == factor_location_2[0] + "区人民":
                    # print("判断错误？")
                    factors_ok.append("解决纠纷-甲方住所地")
                else:
                    factors_error["解决纠纷-甲方住所地"] = "甲方住所地与合同页中描述不符"
                    hp.addRemarkInDoc(word, document, "均可向", "甲方住所地法院与合同页中描述不符")
                    # factors_to_inform['解决纠纷-甲方住所地'] = "请审核甲方住所地是否填写"
            elif flag2:
                # if factor==factor_location_1[0]+"省" or factor==factor_location_2[0]+"市":
                if factor == factor_location_1[0] + "省" or factor == factor_location_1[0] + "省高级" or factor == \
                        factor_location_1[0] + "省高级人民" \
                        or factor == factor_location_2[0] + "市" or factor == factor_location_2[0] + "市中级" or factor == \
                        factor_location_2[0] + "市中级人民" or factor == factor_location_2[0] + "市高级" or factor == \
                        factor_location_2[0] + "市高级人民":
                    factor == factors_ok.append("解决纠纷-甲方住所地")
                else:
                    factors_error["解决纠纷-甲方住所地"] = "甲方住所地与合同页中描述不符"
                    hp.addRemarkInDoc(word, document, "均可向", "甲方住所地法院与合同页中描述不符")
                    # factors_to_inform['解决纠纷-甲方住所地'] = "请审核甲方住所地是否填写"
            elif flag1:
                # if factor==factor_location_1[0]+"省":
                if factor == factor_location_1[0] + "省" or factor == factor_location_1[0] + "省高级" or factor == \
                        factor_location_1[0] + "省高级人民":
                    factors_ok.append("解决纠纷-甲方住所地")
                else:
                    factors_error["解决纠纷-甲方住所地"] = "甲方住所地与合同页中描述不符"
                    hp.addRemarkInDoc(word, document, "均可向", "甲方住所地法院与合同页中描述不符")
                    # factors_to_inform['解决纠纷-甲方住所地'] = "请审核甲方住所地是否填写"
            else:
                factors_error["解决纠纷-甲方住所地"] = "甲方住所地与合同页中描述不符"
                hp.addRemarkInDoc(word, document, "均可向", "甲方住所地法院与合同页中描述不符")
                # factors_to_inform['解决纠纷-甲方住所地'] = "请审核甲方住所地是否填写"
    except:
        # factors_miss.append("甲方住所地法院合同要素缺失")
        factors_miss_block.append(f"\n风险名称：争议解决条款缺失\n风险偏向：(全局)不利于买卖双方\n"
                                  f"风险提示：建议增加以下条款：双方发生纠纷应友好协商解决。协商不成则应提交买方所在地人民法院诉讼解决，本条款在合同终止后仍然有效。")

    # 11.合同份数
    flag = 1
    factors_miss = []
    try:
        match = '本合同一式(.*?)份'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["合同总份数"] = factor
        if factor != "":
            factors_ok.append("合同总份数")
        else:
            factors_error["合同总份数"] = "合同总份数未填写"
            hp.addRemarkInDoc(word, document, "本合同一式", "合同份数未填写")
            # factors_to_inform['合同总份数'] = "请审核合同份数是否填写"
    except:
        factors_miss.append("合同份数合同要素缺失")
        flag = 0

    try:
        match = '份，买方(.*?)份'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["买方合同份数"] = factor
        if factor != "":
            factors_ok.append("买方合同份数")
        else:
            factors_error["买方合同份数"] = "买方合同份数未填写"
            hp.addRemarkInDoc(word, document, "份，买方", "买方合同份数未填写")
            # factors_to_inform['买方合同总份数'] = "请审核买方合同份数是否填写"
    except:
        factors_miss.append("买方合同份数合同要素缺失")
        flag = 0

    try:
        match = '份，卖方(.*?)份'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print(factor)
        factors["卖方份数"] = factor
        if factor != "":
            factors_ok.append("卖方合同份数")
        else:
            factors_error["卖方合同份数"] = "卖方合同份数未填写"
            hp.addRemarkInDoc(word, document, "份，卖方", "卖方合同份数未填写")
            # factors_to_inform['卖方合同总份数'] = "请审核卖方合同份数是否填写"
    except:
        factors_miss.append("卖方合同份数合同要素缺失")
        flag = 0

    if flag == 0:
        factors_miss_block.append(f"\n风险名称：合同生效条款缺失\n风险偏向：(全局)不利于买卖双方\n"
                                  f"风险提示：建议增加以下条款：本合同一式【】份，合同双方各执【】份，均具有同等法律效力。本合同由【双方加盖公章或合同专用章】之日或双方约定的其他生效条件成就时生效。")

    if "合同变更" in text:
        pass
    else:
        factors_miss_block.append(
            f"\n风险名称：合同变更条款缺失\n风险偏向：(全局)不利于买卖双方\n风险提示：建议增加以下条款：如因生产资料、生产设备、生产工艺或市场发生重大变化，卖方须变更货物的品种、花色、规格、质量，卖方应提前【】日与买方协商，双方另行签订书面补充协议。")

    # 将缺失要素加入进列表，在word文档中以批注的形式显示
    '''
    if len(factors_miss) != 0:
        # print("居然不是空的")
        # hp.addRemarkInDoc(word, document, "\n", "合同条款不完整，缺失以下要素：\n")
        str_miss = "\n".join(factors_miss)
        hp.addRemarkInDoc(word, document, "\n", str_miss)
    '''
    # n=1
    # n={ncount}+'.\n'
    if len(factors_miss_block) != 0:
        str_miss = '\n~'.join(factors_miss_block)
        # n+=1
        hp.addRemarkInDoc(word, document, "\n", str_miss)

    # print(factors_miss)
    # print(len(factors_miss))

    try:
        copy_path = processed_file_sava_dir + "/" + filePath.split("/")[-1]
        filePath = hp.str_insert(copy_path, copy_path.index(".doc"), "(已审查)")
        # filePath_1=hp.str_insert(copy_path, copy_path.index(".doc"), "(完整审查)")
        print(filePath)
        document.SaveAs(filePath)
        document.Close()
        factors1, factors_ok1, factors_error1, factors_to_inform1, word = None_standard_contract.Buy_Sell_contract(
            filePath,
            processed_file_sava_dir)
        os.remove(filePath)
        print("中间文件已删除")
        print(factors, factors_ok, factors_error, factors_to_inform, factors_miss)
        # word.Quit()
    except Exception as e:
        print(e)
    # print(factors_to_inform.keys())
    # print(factors_error.keys())
    # print(factors_to_inform["甲方纳税人识别号"])
    # print(len(factors),len(factors_error),len(factors_ok))

    return factors, factors_ok, factors_error, factors_to_inform


# ++++++++++++++++++++++++++++++买卖合同代码结束++++++++++++++++++++++++++++++++++++++++#

# change by suchao
def rent_contract(text, filePath, processed_file_sava_dir):
    return processFuncRent(text, filePath, processed_file_sava_dir)


def construction_contract(text, tables, filePath, vprocessed_file_sava_dir, filePath_zhaobiao=None):
    return processFunc3(tables, text, filePath, vprocessed_file_sava_dir, filePath_zhaobiao)


#  采购合同  change by qy
def purchase_and_warehousing_contract(text, tables, filePath, processed_file_sava_dir):
    return processFunc(text, tables, filePath, processed_file_sava_dir)


# 物业管理合同，by wzk
def property_management_contract(text, tables, filePath, processed_file_sava_dir):
    try:
        pythoncom.CoInitialize()
        word = Dispatch('Word.Application')
        pythoncom.CoInitialize()
        word.Documents.close()
        word.Quit()
        word.Visible = 0  # 后台打开word文档
    except Exception as ex:
        print(ex)
    try:
        document = word.Documents.Open(FileName=filePath)
    except Exception as ex:
        print(ex)
    factors = {}
    factors_ok = []
    factors_error = {}
    factors_to_inform = {}
    factors_miss = []

    # 用于标志每个大项是否有缺失
    miss_flag = False
    # print(len(tables))
    # for i in range(0, len(tables)):
    #     for row in range(0,len(tables[i].rows)):
    #         for col in range(0,len(tables[i].columns)):
    #             print(tables[i].cell(row, col).text)
    table = tables[1]
    for i in range(0, len(table.rows)):
        for j in range(0, len(table.columns)):
            cell = table.cell(i, j)
            print(f"单元格（{i}，{j}）的值为：{cell.text}\n")
            print(cell.text)
    standard = ['劳动纪律', '巡查管理', '员工培训', '物资管理', '办公区域', '卫生间', '会议室及观景阳台', '天花板', '楼道', '电梯', '室外道路', '园区垃圾箱', '其他',
                '值班时间', '外来人员车辆', '值班室清洁', '安全检查', '前台接待', '登记', '会务服务', '报刊分发', '邮寄', '文印', '食品安全', '安全记录', '食堂卫生',
                '维修处理', '日常巡查', '绿化补种', '病虫害防治', '除草、施肥']

    # 检查考核标准表
    check_standard = []

    # 主体审查
    try:
        match = '甲方：【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['甲方'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('甲方')
        else:
            factors_error['甲方'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "甲方", "【】中应填写有或无")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    # 审查【】之外的内容
    try:
        match = '甲方：(.*?)\n'
        factor = re.findall(match, text)[0].replace(' ', '').replace('【', '').replace('】', '').replace('有', '').replace(
            '无', '')
        factors['甲方'] = factor
        if factor != '':
            factors_ok.append('甲方1')
        else:
            factors_error['甲方1'] = "甲方未填写"
            hp.addRemarkInDoc(word, document, "甲方", "甲方未填写")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    try:
        match = '住所地：【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['甲方住所地'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('甲方住所地')
        else:
            factors_error['甲方住所地'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "住所地", "【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    try:
        match = '住所地：(.*?)\n'
        factor = re.findall(match, text)[0].replace(' ', '').replace('【', '').replace('】', '').replace('有', '').replace(
            '无', '')
        factors['甲方住所地1'] = factor
        if factor != '':
            factors_ok.append('甲方住所地1')
        else:
            factors_error['甲方住所地1'] = "甲方住所地未填写"
            hp.addRemarkInDoc(word, document, "住所地", "甲方住所地未填写")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    try:
        match = '法定代表人/负责人：【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['甲方法定代表人'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('甲方法定代表人')
        else:
            factors_error['甲方法定代表人'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "法定代表人/负责人", "【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    try:
        match = '法定代表人/负责人：(.*?)\n'
        factor_corporation_1 = re.findall(match, text)[0].replace(' ', '').replace('【', '').replace('】', '').replace(
            '有', '').replace(
            '无', '')
        factors['甲方法定代表人1'] = factor_corporation_1
        if factor_corporation_1 != '':
            factors_ok.append('甲方法定代表人1')
        else:
            factors_error['甲方法定代表人1'] = "甲方法定代表人未填写"
            hp.addRemarkInDoc(word, document, "法定代表人/负责人", "甲方法定代表人/负责人未填写")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    try:
        match = '组织信用代码/身份证号：【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['甲方组织信用代码/身份证号'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('甲方组织信用代码/身份证号')
        else:
            factors_error['甲方组织信用代码/身份证号'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "组织信用代码/身份证号", "【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    try:
        match = '组织信用代码/身份证号：(.*?)\n'
        factor = re.findall(match, text)[0].replace(' ', '').replace('【', '').replace('】', '').replace('有', '').replace(
            '无', '')
        factors['甲方组织信用代码/身份证号1'] = factor
        if factor != "":
            if UnifiedSocialCreditIdentifier().check_code(factor, 'sc'):
                factors_ok.append("甲方组织信用代码/身份证号")
            else:

                if checkIdCard(factor) == 'ok':
                    factors_ok.append("甲方组织信用代码/身份证号1")
                else:
                    factors_error["甲方组织信用代码/身份证号1"] = "组织信用代码未填写正确或" + checkIdCard(factor)
                    hp.addRemarkInDoc(word, document, "组织信用代码/身份证号", "甲方组织信用代码/身份证号校验未通过")
                    # factors_to_inform['甲方（买方）统一社会信用代码/身份证号码审核提示'] = "请审核甲方（买方）统一社会信用代码/身份证号码是否填写正确"
        else:
            factors_error["甲方组织信用代码/身份证号1"] = "组织信用代码/身份证号未填写完整"
            # factors_to_inform['甲方（买方）统一社会信用代码/身份证号码审核提示'] = "请审核甲方（买方）统一社会信用代码/身份证号码是否填写正确"
            hp.addRemarkInDoc(word, document, "组织信用代码/身份证号", "甲方组织信用代码/身份证号未填写")

    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    try:
        match = '联系方式：【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['甲方联系方式'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('甲方联系方式')
        else:
            factors_error['甲方联系方式'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "联系方式", "【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    try:
        match = '联系方式：(.*?)\n'
        factor = re.findall(match, text)[0].replace(' ', '').replace('【', '').replace('】', '').replace('有', '').replace(
            '无', '')
        factors['甲方联系方式1'] = factor
        if factor != "":
            factor = factor.replace("（", "").replace("）", "").replace("-", "")
            # if len(factor)==10:
            # factors_ok.append("甲方（买方）联系方式")
            # hp.addRemarkInDoc(word, document, "联系方式", "联系方式为10位号码，请人工检查")
            # else:
            if isTelPhoneNumber(factor) != "Error":
                factors_ok.append("甲方联系方式1")
            else:
                factors_error["甲方联系方式1"] = "甲方联系方式填写有误"
                # factors_to_inform['甲方（买方）联系方式审核提示'] = "请审核甲方（买方）联系方式是否填写正确"
                hp.addRemarkInDoc(word, document, "联系方式", "甲方联系方式错误")
        else:
            factors_error["甲方联系方式1"] = "甲方联系方式不完整"
            # factors_to_inform['甲方（买方）联系方式审核提示'] = "请审核甲方（买方）联系方式是否填写正确"
            hp.addRemarkInDoc(word, document, "联系方式", "甲方联系方式未填写完整")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    # 要素缺失批注，整个甲方主体信息一起提示，模块提示
    if miss_flag:
        factors_miss.append("甲方主体信息要素缺失\n")

    miss_flag = False

    try:
        match = '乙方：【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['乙方'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('乙方')
        else:
            factors_error['乙方'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "乙方", "【】中应填写有或无")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    # 审查【】之外的内容
    try:
        match = '乙方：(.*?)\n'
        factor = re.findall(match, text)[0].replace(' ', '').replace('【', '').replace('】', '').replace('有', '').replace(
            '无', '')
        factors['乙方'] = factor
        if factor != '':
            factors_ok.append('乙方1')
        else:
            factors_error['乙方1'] = "甲方未填写"
            hp.addRemarkInDoc(word, document, "乙方", "乙方未填写")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    try:
        match = '住所地：【(.*?)】'
        factor = re.findall(match, text)[1].replace(' ', '')
        factors['乙方住所地'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('乙方住所地')
        else:
            factors_error['乙方住所地'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "乙方", "住所地【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    try:
        match = '住所地：(.*?)\n'
        factor = re.findall(match, text)[1].replace(' ', '').replace('【', '').replace('】', '').replace('有', '').replace(
            '无', '')
        factors['乙方住所地1'] = factor
        if factor != '':
            factors_ok.append('乙方住所地1')
        else:
            factors_error['乙方住所地1'] = "乙方住所地未填写"
            hp.addRemarkInDoc(word, document, "乙方", "乙方住所地未填写")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    try:
        match = '法定代表人/负责人：【(.*?)】'
        factor = re.findall(match, text)[1].replace(' ', '')
        factors['乙方法定代表人'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('乙方法定代表人')
        else:
            factors_error['乙方法定代表人'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "乙方", "法定代表人【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    try:
        match = '法定代表人/负责人：(.*?)\n'
        factor_corporation_2 = re.findall(match, text)[1].replace(' ', '').replace('【', '').replace('】', '').replace(
            '有', '').replace(
            '无', '')
        factors['乙方法定代表人1'] = factor_corporation_2
        if factor_corporation_2 != '':
            factors_ok.append('乙方法定代表人1')
        else:
            factors_error['乙方法定代表人1'] = "乙方法定代表人未填写"
            hp.addRemarkInDoc(word, document, "乙方", "乙方法定代表人/负责人未填写")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    try:
        match = '组织信用代码/身份证号：【(.*?)】'
        factor = re.findall(match, text)[1].replace(' ', '')
        factors['乙方组织信用代码/身份证号'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('乙方组织信用代码/身份证号')
        else:
            factors_error['乙方组织信用代码/身份证号'] = "组织信用代码，身份证号【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "乙方", "【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    try:
        match = '组织信用代码/身份证号：(.*?)\n'
        factor = re.findall(match, text)[1].replace(' ', '').replace('【', '').replace('】', '').replace('有', '').replace(
            '无', '')
        factors['乙方组织信用代码/身份证号1'] = factor
        if factor != "":
            if UnifiedSocialCreditIdentifier().check_code(factor, 'sc'):
                factors_ok.append("乙方组织信用代码/身份证号")
            else:

                if checkIdCard(factor) == 'ok':
                    factors_ok.append("乙方组织信用代码/身份证号1")
                else:
                    factors_error["乙方组织信用代码/身份证号1"] = "组织信用代码未填写正确或" + checkIdCard(factor)
                    hp.addRemarkInDoc(word, document, "乙方", "乙方组织信用代码/身份证号校验未通过")
                    # factors_to_inform['甲方（买方）统一社会信用代码/身份证号码审核提示'] = "请审核甲方（买方）统一社会信用代码/身份证号码是否填写正确"
        else:
            factors_error["乙方组织信用代码/身份证号1"] = "组织信用代码/身份证号未填写完整"
            # factors_to_inform['甲方（买方）统一社会信用代码/身份证号码审核提示'] = "请审核甲方（买方）统一社会信用代码/身份证号码是否填写正确"
            hp.addRemarkInDoc(word, document, "乙方", "甲方组织信用代码/身份证号未填写")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    try:
        match = '联系方式：【(.*?)】'
        factor = re.findall(match, text)[1].replace(' ', '')
        factors['乙方联系方式'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('乙方联系方式')
        else:
            factors_error['乙方联系方式'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "乙方", "联系方式【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    try:
        match = '联系方式：(.*?)\n'
        factor = re.findall(match, text)[1].replace(' ', '').replace('【', '').replace('】', '').replace('有', '').replace(
            '无', '')
        factors['乙方联系方式1'] = factor
        if factor != "":
            factor = factor.replace("（", "").replace("）", "").replace("-", "")
            # if len(factor)==10:
            # factors_ok.append("甲方（买方）联系方式")
            # hp.addRemarkInDoc(word, document, "联系方式", "联系方式为10位号码，请人工检查")
            # else:
            if isTelPhoneNumber(factor) != "Error":
                factors_ok.append("乙方联系方式1")
            else:
                factors_error["乙方联系方式1"] = "乙方联系方式填写有误"
                # factors_to_inform['甲方（买方）联系方式审核提示'] = "请审核甲方（买方）联系方式是否填写正确"
                hp.addRemarkInDoc(word, document, "乙方", "乙方联系方式错误")
        else:
            factors_error["乙方联系方式1"] = "乙方联系方式不完整"
            # factors_to_inform['甲方（买方）联系方式审核提示'] = "请审核甲方（买方）联系方式是否填写正确"
            hp.addRemarkInDoc(word, document, "乙方", "乙方联系方式未填写完整")
    except:
        # factors_miss.append("甲方缺失")
        miss_flag = True

    # 要素缺失批注，整个甲方主体信息一起提示，模块提示
    if miss_flag:
        factors_miss.append("乙方主体信息要素缺失\n")

    miss_flag = False

    try:
        match = '物业服务范围【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['物业服务范围'] = factor
        if factor != '':
            factors_to_inform['物业服务范围'] = '请审核物业服务范围是否与招标文件一致'
            hp.addRemarkInDoc(word, document, '物业服务范围', "请审核物业服务范围是否与招标文件一致")
        else:
            factors_error['物业服务范围'] = '物业服务范围未填写'
            hp.addRemarkInDoc(word, document, '物业服务范围', "物业服务范围未填写")
    except:
        miss_flag = True

    try:
        match = '物业服务地点【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['物业服务地点'] = factor
        if factor != '':
            factors_to_inform['物业服务地点'] = '请审核物业服务地点是否与招标文件一致'
            hp.addRemarkInDoc(word, document, '物业服务地点', "请审核物业服务地点是否与招标文件一致")
        else:
            factors_error['物业服务地点'] = '物业服务地点未填写'
            hp.addRemarkInDoc(word, document, '物业服务地点', "物业服务地点未填写")
    except:
        miss_flag = True

    if miss_flag:
        factors_miss.append('物业服务概况要素不完整\n')

    miss_flag = False

    try:

        match = '本物业服务合同须期限壹年，自【(.*?)】年【(.*?)】月【(.*?)】日至【(.*?)】年【(.*?)】月【(.*?)】日'
        factor = list(tuple(re.findall(match, text)[0]))
        print(factor)
        print(len(factor))
        factor = [i for i in factor if i != '']
        factors['合同期限'] = factor
        print(factor)
        print(len(factor))
        if len(factor) == 6:
            for i in range(6):
                factor[i] = factor[i].replace(' ', '')
            factors["合同期限"] = f'{factor[0]}-{factor[1]}-{factor[2]}至{factor[3]}-{factor[4]}-{factor[5]}'
            if isRightDate(factor[0], factor[1], factor[2]) and isRightDate(factor[3], factor[4], factor[5]):
                # factors_ok.append("三1&协议期限")
                day_start = datetime.date(int(factor[0]), int(factor[1]), int(factor[2]))
                day_end = datetime.date(int(factor[3]), int(factor[4]), int(factor[5]))
                period = day_end.__sub__(day_start).days
                if period == 365 or period == 366 or period == 364:
                    factors_ok.append('合同期限')
                elif period < 365:
                    factors_error['合同期限'] = "合同期限错误，期限未达一年"
                    hp.addRemarkInDoc(word, document, '本物业服务合同须期限', '合同期限填写错误，期限未达一年')
                else:
                    factors_error['合同期限'] = "合同期限错误，期限超过一年"
                    hp.addRemarkInDoc(word, document, '本物业服务合同须期限', '合同期限填写错误，期限超过一年')

            else:
                factors_error["合同期限"] = "合同期限时间未填写规范"
                addRemarkInDoc(word, document, "本物业服务合同须期限", "合同期限时间未填写规范")
        else:
            factors_error["合同期限"] = "合同期限时间未填写完整"
            addRemarkInDoc(word, document, "本物业服务合同须期限", "合同期限时间未填写完整")
    except:
        miss_flag = True

    if miss_flag:
        factors_miss.append('合同期限要素缺失\n')

    miss_flag = False

    try:
        match = '安全保卫【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['安全保卫'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('安全保卫')
        else:
            factors_error['安全保卫'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "安全保卫", "安全保卫【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('安全保卫要素缺失\n')

    miss_flag = False

    try:
        match = '卫生保洁【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['卫生保洁'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('卫生保洁')
        else:
            factors_error['卫生保洁'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "卫生保洁", "卫生保洁【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('卫生保洁要素缺失\n')

    miss_flag = False

    try:
        match = '（11）其他：(.*?)\n'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['卫生保洁-其他'] = factor
        if factor != '':
            factors_ok.append('卫生保洁-其他')
        else:
            factors_error['卫生保洁-其他'] = "应填写无"
            hp.addRemarkInDoc(word, document, "（11）其他：", "如未填写内容，应填写\"无\"")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('卫生保洁-其他要素缺失\n')

    miss_flag = False

    try:
        match = '接待及会务服务【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['接待及会务服务'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('接待及会务服务')
        else:
            factors_error['接待及会务服务'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "接待及会务服务", "接待及会务服务【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('接待及会务服务要素缺失\n')

    miss_flag = False

    try:
        match = '员工食堂【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['员工食堂'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('员工食堂')
        else:
            factors_error['员工食堂'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "员工食堂", "员工食堂【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('员工食堂要素缺失\n')

    miss_flag = False

    try:
        match = '设施设备维修【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['设施设备维修'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('设施设备维修')
        else:
            factors_error['设施设备维修'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "设施设备维修", "设施设备维修【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('设施设备维修要素缺失\n')

    miss_flag = False

    try:
        match = '绿化养护【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['绿化养护'] = factor
        if factor != '' and (factor == '有' or factor == '无'):
            factors_ok.append('绿化养护')
        else:
            factors_error['绿化养护'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "绿化养护", "绿化养护【】中应填写有或无")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('绿化养护要素缺失\n')

    miss_flag = False

    try:
        match = '人员配备要求(.*?)'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['人员配备要求'] = factor
        factors_to_inform['人员配备要求'] = "请核实人员配备要求与招投标文件相关内容一致"

        hp.addRemarkInDoc(word, document, "人员配备要求", "请核实人员配备要求与招投标文件相关内容一致")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('人员配备要求表格要素缺失\n')

    miss_flag = False

    try:
        match_total = '甲方每年向乙方支付物业管理服务费为【(.*?)】元'
        match_total_kanji = '大写：【(.*?)】'  # 价格中文大写，用于检查大小写是否一致
        factor_total = re.findall(match_total, text)[0].replace(" ", "")

        factor_total_kanji = re.findall(match_total_kanji, text)[0]
        # 去掉圆整等汉字
        factor_total_kanji = factor_total_kanji.replace(" ", "").replace("整", "").replace("圆", "").replace("元", "")
        # print(factor_total_kanji)
        # print(type(factor_total_kanji))
        factors["货品总金额（价税合计）"] = factor_total
        # print(factor_total)
        if factor_total == '':
            # print("货品总金额根本没匹配到")
            factors_error['物业管理费'] = "物业管理费未填写"
            hp.addRemarkInDoc(word, document, "甲方每年向乙方支付物业管理服务费为", "物业管理费未填写")
            # factors_to_inform['货品总金额'] = "请审核货品总金额是否填写"
        else:
            container = hp.money_en_to_cn(float(factor_total)).replace("圆", "")
            # print(type(container))
            if container == factor_total_kanji:
                # 金额大小写能匹配
                factors_ok.append("物业管理费")
            else:
                factors_error['物业管理费'] = "物业管理费人民币大写与数值不一致"
                hp.addRemarkInDoc(word, document, "甲方每年向乙方支付物业管理服务费为", "物业管理费人民币大写与数值不一致")
                # factors_to_inform['货品总金额'] = "请审核货品总金额大小写是否一致"
    except:
        # factors_miss.append("货品总金额（价税合计）合同要素缺失")
        miss_flag = True

    try:
        match = '物管费按【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['物管费支付'] = factor
        if factor != '' and (factor == '月' or factor == '季度'):
            factors_ok.append('物管费支付')
        else:
            factors_error['物管费支付'] = "【】中应填写有或无"
            hp.addRemarkInDoc(word, document, "物管费按", "【】中应填写月或季度")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    try:
        match = '甲方在支付月【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['物管费支付期限'] = factor
        if factor != '' and int(factor) <= 31:
            factors_ok.append('物管费支付期限')
            factors_to_inform['物管费支付期限'] = "请向财务部核实该支付日期"
            hp.addRemarkInDoc(word, document, "甲方在支付月", "请向财务部核实该支付日期")
        else:
            factors_error['物管费支付期限'] = "物管费支付期限未填写或错误"
            hp.addRemarkInDoc(word, document, "甲方在支付月", "物管费支付期限未填写或填写错误")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('第五条费用要素不完整\n')

    miss_flag = False

    try:
        match_total = '本合同履约保证金金额为【(.*?)】元'
        match_total_kanji = '大写：【(.*?)】'  # 价格中文大写，用于检查大小写是否一致
        factor_total = re.findall(match_total, text)[0].replace(" ", "")

        factor_total_kanji = re.findall(match_total_kanji, text)[1]
        # 去掉圆整等汉字
        factor_total_kanji = factor_total_kanji.replace(" ", "").replace("整", "").replace("圆", "").replace("元", "")
        # print(factor_total_kanji)
        # print(type(factor_total_kanji))
        factors["履约保证金"] = factor_total
        # print(factor_total)
        if factor_total == '':
            # print("货品总金额根本没匹配到")
            factors_error['履约保证金'] = "履约保证金未填写"
            hp.addRemarkInDoc(word, document, "本合同履约保证金金额为", "履约保证金未填写")
            # factors_to_inform['货品总金额'] = "请审核货品总金额是否填写"
        else:
            container = hp.money_en_to_cn(float(factor_total)).replace("圆", "")
            # print(type(container))
            if container == factor_total_kanji:
                # 金额大小写能匹配
                factors_ok.append("履约保证金")
            else:
                factors_error['履约保证金'] = "履约保证金人民币大写与数值不一致"
                hp.addRemarkInDoc(word, document, "本合同履约保证金金额为", "履约保证金人民币大写与数值不一致")
                # factors_to_inform['货品总金额'] = "请审核货品总金额大小写是否一致"
    except:
        # factors_miss.append("货品总金额（价税合计）合同要素缺失")
        miss_flag = True

    try:

        match = '应于【(.*?)】年【(.*?)】月【(.*?)】日前一次性支付'
        factor = list(tuple(re.findall(match, text)[0]))
        # print(factor)
        # print(len(factor))
        factor = [i for i in factor if i != '']
        # print(factor)
        # print(len(factor))
        factors['履约保证金支付期限'] = factor
        if len(factor) == 3:
            for i in range(3):
                factor[i] = factor[i].replace(' ', '')
            factors["履约保证金支付期限"] = f'{factor[0]}-{factor[1]}-{factor[2]}'
            if isRightDate(factor[0], factor[1], factor[2]):
                factors_ok.append('履约保证金支付期限')
            else:
                factors_error['履约保证金支付期限'] = "履约保证金支付期限日期填写错误"
                hp.addRemarkInDoc(word, document, "日前一次性支付", "履约保证金支付期限日期填写不符合规范")

        else:
            factors_error["履约保证金支付期限"] = "履约保证金支付期限未填写完整"
            addRemarkInDoc(word, document, "日前一次性支付", "履约保证金支付期限日期未填写完整")
    except:
        miss_flag = True

    if miss_flag:
        factors_miss.append("第六条履约保证金要素不完整\n")

    miss_flag = False

    try:
        match = '附件：(.*?)\n'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['附件'] = factor
        if factor != '':
            factors_ok.append('附件')
            factors_to_inform['附件'] = "请向主管部门核实附件一的内容"
            hp.addRemarkInDoc(word, document, "对物管工作进行定期或不定期的单方检查考核", "请向主管部门核实附件一的内容")
            # hp.addRemarkInDoc(word, document, "附件：", "请向主管部门核实附件一的内容")
        else:
            factors_error['附件'] = "附件缺失"
            hp.addRemarkInDoc(word, document, "对物管工作进行定期或不定期的单方检查考核", "审查考核附件缺失")
        # 考核标准表格有问题，格式不对，只能读出表头，用其他表格测试可以读取
        if len(tables) == 3:
            # 审查第三个表，即考核标准表
            print('111111')
            table_standard = tables[2]
            print(type(table_standard.cell(1, 1).text))
            for i in range(1, len(table_standard.rows)):
                info = table_standard.cell(i, 1).text.replace(' ', '')

                check_standard.append(info)
            print(check_standard)
            print(set(check_standard).issubset(set(standard)))
            print(set(check_standard) == (set(standard)))
            if set(check_standard).issubset(set(standard)) or set(check_standard) == (set(standard)):
                factors_ok.append("考核标准表")
                hp.addRemarkInDoc(word, document, "附件：", "具体的扣分项及考核标准请使用单位具体细化完善")
            else:
                factors_error['考核标准表'] = "考核标准表不完整"
                hp.addRemarkInDoc(word, document, "附件：", "请核实考核办法及标准表是否包含了所有要素，具体的扣分项及考核标准请使用单位具体细化完善")
        else:
            factors_miss.append("考核标准表缺失\n")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True
    print(standard)
    try:
        match = '如在【(.*?)】日'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['考核期限'] = factor
        if factor != '' and int(factor) <= 7 and int(factor) >= 1:
            factors_ok.append('考核期限')
            # factors_to_inform['考核期限'] = "请向财务部核实该支付日期"
            # hp.addRemarkInDoc(word,document,"甲方在支付月","请向财务部核实该支付日期")
        else:
            factors_error['考核期限'] = "考核期限未填写或错误，该数值在1-7之间选择"
            hp.addRemarkInDoc(word, document, "如在【", "考核期限未填写或错误，该数值在1-7之间选择")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('第七条检查考核要素不完整\n')

    miss_flag = False

    try:
        match = '服务费总额的【(.*?)】'
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['违约金'] = factor
        if factor != '' and int(factor) <= 30 and int(factor) >= 10:
            factors_ok.append('违约金')
            # factors_to_inform['考核期限'] = "请向财务部核实该支付日期"
            # hp.addRemarkInDoc(word,document,"甲方在支付月","请向财务部核实该支付日期")
        else:
            factors_error['违约金'] = "违约金未填写或错误，该数值在10-30之间"
            hp.addRemarkInDoc(word, document, "服务费总额的", "违约金未填写或错误，该数值在10-30之间")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    if miss_flag:
        factors_miss.append('第十条违约责任违约金要素缺失\n')

    miss_flag = False

    try:
        match = '本合同一式【(.*?)】份'
        print('1')
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['合同总份数'] = factor
        if factor != '':
            factors_ok.append('总份数')
            # factors_to_inform['附件'] = "请向主管部门核实附件一的内容"
            # hp.addRemarkInDoc(word,document,"本合同一式","请向主管部门核实附件一的内容")
        else:
            factors_error['总份数'] = "合同总份数未填写"
            hp.addRemarkInDoc(word, document, "本合同一式", "合同总份数未填写")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True
    print(miss_flag)

    try:
        match = '甲乙双方各执【(.*?)】份'
        print('2')
        factor = re.findall(match, text)[0].replace(' ', '')
        factors['合同双方份数'] = factor
        if factor != '':
            factors_ok.append('双方份数')
            # factors_to_inform['附件'] = "请向主管部门核实附件一的内容"
            # hp.addRemarkInDoc(word,document,"本合同一式","请向主管部门核实附件一的内容")
        else:
            factors_error['双方份数'] = "合同双方份数未填写"
            hp.addRemarkInDoc(word, document, "甲乙双方各执", "合同双方份数未填写")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True
    print(miss_flag)

    table = tables[1]
    content1 = table.cell(0, 0).text.replace(' ', '')
    print(content1)
    print(type(content1))
    try:
        match = '甲方：(.*)'

        factor = re.findall(match, content1)[0].replace(' ', '')
        print(factor)
        factors['甲方-尾部'] = factor

        print('甲方')
        if factor != '':
            factors_ok.append('甲方-尾部')
            # factors_to_inform['附件'] = "请向主管部门核实附件一的内容"
            # hp.addRemarkInDoc(word,document,"本合同一式","请向主管部门核实附件一的内容")
        else:
            factors_error['甲方-尾部'] = "甲方-尾部未填写"
            hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "甲方未填写")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    content1 = table.cell(0, 1).text
    print(content1)
    print(type(content1))
    try:
        match = '乙方：(.*)'

        factor = re.findall(match, content1)[0].replace(' ', '')
        factors['乙方-尾部'] = factor

        print('乙方')
        if factor != '':
            factors_ok.append('乙方-尾部')
            # factors_to_inform['附件'] = "请向主管部门核实附件一的内容"
            # hp.addRemarkInDoc(word,document,"本合同一式","请向主管部门核实附件一的内容")
        else:
            factors_error['乙方-尾部'] = "乙方-尾部未填写"
            hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "乙方未填写")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True

    content1 = table.cell(2, 0).text
    print(content1)
    try:
        match = '法定（授权）代表人：(.*)'
        # print("法定代表人1")
        factor = re.findall(match, content1)[0].replace(' ', '')
        print(factor)
        factors['甲方法定（授权）代表人'] = factor
        if factor != '':
            if factor == factor_corporation_1:
                factors_ok.append('甲方法定（授权）代表人')
            else:
                factors_error['甲方法定（授权）代表人'] = '甲方法定（授权）代表人与合同首部主体信息不一致，请核实'
                hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", '甲方法定（授权）代表人与合同首部主体信息不一致，请核实')
        else:
            factors_error['甲方法定（授权）代表人'] = "甲方法定（授权）代表人未填写"
            hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "甲方法定（授权）代表人未填写")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True
    print(miss_flag)
    content1 = table.cell(2, 1).text
    print(content1)
    try:
        match = '法定（授权）代表人：(.*)'

        factor = re.findall(match, content1)[0].replace(' ', '')
        print(factor)
        factors['乙方法定（授权）代表人'] = factor
        if factor != '':
            if factor == factor_corporation_2:
                factors_ok.append('乙方法定（授权）代表人')
            else:
                factors_error['乙方法定（授权）代表人'] = '乙方法定（授权）代表人与合同首部主体信息不一致，请核实'
                hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", '乙方法定（授权）代表人与合同首部主体信息不一致，请核实')
        else:
            factors_error['乙方法定（授权）代表人'] = "乙方法定（授权）代表人未填写"
            hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "乙方法定（授权）代表人未填写")
    except:
        # factors_miss.append("甲方住所地缺失")
        miss_flag = True
    print(miss_flag)

    content1 = table.cell(3, 0).text
    print(content1)

    try:

        match = '(.*?)年(.*?)月(.*?)日'
        factor = list(tuple(re.findall(match, content1)[0]))
        # print(factor)
        # print(len(factor))
        factor = [i for i in factor if i != '']
        factors['签字时间甲方'] = factor
        # print(factor)
        # print(len(factor))
        if len(factor) == 3:
            for i in range(3):
                factor[i] = factor[i].replace(' ', '')
            factors["签字时间甲方"] = f'{factor[0]}-{factor[1]}-{factor[2]}'
            if isRightDate(factor[0], factor[1], factor[2]):
                factors_ok.append('签字期限甲方')
            else:
                factors_error['签字时间甲方'] = "甲方签字日期填写错误"
                hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "甲方签字时间填写错误")

        else:
            factors_error["签字时间甲方"] = "签字时间甲方未填写完整"
            addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "甲方签字时间未填写完整")
    except:
        miss_flag = True

    content1 = table.cell(3, 1).text
    print(content1)
    try:

        match = '(.*?)年(.*?)月(.*?)日'
        factor = list(tuple(re.findall(match, content1)[0]))
        # print(factor)
        # print(len(factor))
        factor = [i for i in factor if i != '']
        factors['签字时间乙方'] = factor
        # print(factor)
        # print(len(factor))
        if len(factor) == 3:
            for i in range(3):
                factor[i] = factor[i].replace(' ', '')
            factors["签字时间乙方"] = f'{factor[0]}-{factor[1]}-{factor[2]}'
            if isRightDate(factor[0], factor[1], factor[2]):
                factors_ok.append('签字期限乙方')
            else:
                factors_error['签字时间乙方'] = "乙方签字日期填写错误"
                hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "乙方签字时间填写错误")

        else:
            factors_error["签字时间乙方"] = "签字时间乙方未填写完整"
            addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "乙方签字时间未填写完整")
    except:
        miss_flag = True

    # table=tables[1]
    # for i in range(0,len(table.rows))
    # try:
    #     match = '甲方：(.*?)\n'
    #
    #     factor = re.findall(match, text)[1].replace(' ', '')
    #     factors['甲方-尾部'] = factor
    #
    #     print('甲方')
    #     if factor != '':
    #         factors_ok.append('甲方-尾部')
    #         # factors_to_inform['附件'] = "请向主管部门核实附件一的内容"
    #         # hp.addRemarkInDoc(word,document,"本合同一式","请向主管部门核实附件一的内容")
    #     else:
    #         factors_error['甲方-尾部'] = "甲方-尾部未填写"
    #         hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "甲方未填写")
    # except:
    #     # factors_miss.append("甲方住所地缺失")
    #     miss_flag = True
    # print(miss_flag)
    # try:
    #     match = '乙方：(.*?)\n'
    #
    #     factor = re.findall(match, text)[1].replace(' ', '')
    #     factors['乙方-尾部'] = factor
    #     print('乙方')
    #     if factor != '':
    #         factors_ok.append('乙方-尾部')
    #         # factors_to_inform['附件'] = "请向主管部门核实附件一的内容"
    #         # hp.addRemarkInDoc(word,document,"本合同一式","请向主管部门核实附件一的内容")
    #     else:
    #         factors_error['乙方-尾部'] = "乙方-尾部未填写"
    #         hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "乙方未填写")
    # except:
    #     # factors_miss.append("甲方住所地缺失")
    #     miss_flag = True
    # print(miss_flag)
    # try:
    #     match = '法定（授权）代表人：(.*?)\n'
    #     print("法定代表人1")
    #     factor = re.findall(match, text)[0].replace(' ', '')
    #     factors['甲方法定（授权）代表人'] = factor
    #     if factor != '':
    #         if factor == factor_corporation_1:
    #             factors_ok.append('甲方法定（授权）代表人')
    #         else:
    #             factors_error['甲方法定（授权）代表人'] = '甲方法定（授权）代表人与合同首部主体信息不一致，请核实'
    #             hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", '甲方法定（授权）代表人与合同首部主体信息不一致，请核实')
    #     else:
    #         factors_error['甲方法定（授权）代表人'] = "甲方法定（授权）代表人未填写"
    #         hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "甲方法定（授权）代表人未填写")
    # except:
    #     # factors_miss.append("甲方住所地缺失")
    #     miss_flag = True
    # print(miss_flag)
    # try:
    #     match = '法定（授权）代表人：(.*?)\n'
    #     print("法定代表人2")
    #     factor = re.findall(match, text)[1].replace(' ', '')
    #     factors['乙方法定（授权）代表人'] = factor
    #     if factor != '':
    #         if factor == factor_corporation_2:
    #             factors_ok.append('乙方法定（授权）代表人')
    #         else:
    #             factors_error['乙方法定（授权）代表人'] = '乙方法定（授权）代表人与合同首部主体信息不一致，请核实'
    #             hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", '乙方法定（授权）代表人与合同首部主体信息不一致，请核实')
    #     else:
    #         factors_error['乙方法定（授权）代表人'] = "乙方法定（授权）代表人未填写"
    #         hp.addRemarkInDoc(word, document, "本合同经双方法定代表人或授权代表签字，并加盖公章或合同专用章后生效", "乙方法定（授权）代表人未填写")
    # except:
    #     # factors_miss.append("甲方住所地缺失")
    #     miss_flag = True
    # print(miss_flag)
    #

    if miss_flag:
        factors_miss.append("第十二条生效及其他要素缺失\n")

    miss_flag = False
    '''
    standard=['劳动纪律']
    check_standard=[]
    if len(tables)==2:
        # 审查第二个表，即考核标准表
        table_standard=tables[1]
        for i in range(1,len(table_standard.rows)):
            info=table_standard.cell(i,1)
            print(info)
            check_standard.append(info)
        if set(check_standard).issubset(standard):
            factors_ok.append("考核标准表")
        else:
            factors_error['考核标准表']="考核标准表不完整"
            hp.addRemarkInDoc(word,document,"附件：","请核实是否考核办法及标准表是否包含了所有要素，具体的扣分项及考核标准请使用单位具体细化完善")
    '''
    try:
        if len(factors_miss) != 0:
            str_miss = ''.join(factors_miss)
            hp.addRemarkInDoc(word, document, "", str_miss)
        copy_path = processed_file_sava_dir + "/" + filePath.split("/")[-1]
        filePath = str_insert(copy_path, copy_path.index(".doc"), "(已审查)")
        print(filePath)
        document.SaveAs(filePath)
        document.Close()
        factors1, factors_ok1, factors_error1, factors_to_inform1, word = None_standard_contract.property_management_contract(
            filePath,
            processed_file_sava_dir)
        os.remove(filePath)
        # word.Quit()
    except Exception as ex:
        print(ex)
    print(factors, factors_ok, factors_error, factors_to_inform)

    return factors, factors_ok, factors_error, factors_to_inform
