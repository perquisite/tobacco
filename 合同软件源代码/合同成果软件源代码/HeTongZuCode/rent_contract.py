import os
import re

import pythoncom

import None_standard_contract
from utils import UnifiedSocialCreditIdentifier, checkIdCard, isTelPhoneNumber, isRightDate, checkQQ, checkEmail, \
    str_insert, addRemarkInDoc, digital_to_Upper, is_contain_dot, isEmail, checkEntersAndSpace
from win32com.client import Dispatch
import helpful as hp


# add by sc

def processFuncRent(text, filePath, processed_file_sava_dir):
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
        document = word.Documents.Open(FileName=filePath, )
    except Exception as ex:
        print(ex)
    factors = {}
    factors_ok = []
    factors_error = {}
    factors_to_inform = {}
    # print(text, tables)

    # 缺失的要素
    missObject = ""

    # 出租方主体审查
    try:
        match = '出租方：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&出租方"] = factor
        if factor != "":
            factors_ok.append("主体&出租方")
        else:
            factors_error["主体&出租方"] = "出租方未填写完整"
            factors_to_inform["主体&出租方"] = "出租方未填写完整"
            addRemarkInDoc(word, document, "出租方", f"要素填写错误：出租方未填写完整")
    except:
        missObject += "要素“出租方”缺失\n"

    try:
        match = '法定代表人/负责人：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&出租方法定代表人/负责人"] = factor
        if factor != "":
            factors_ok.append("主体&出租方法定代表人/负责人")
        else:
            factors_error["主体&出租方法定代表人/负责人"] = "法定代表人/负责人未填写完整"
            factors_to_inform["主体&出租方法定代表人/负责人"] = "出租方法定代表人/负责人未填写完整"
            addRemarkInDoc(word, document, "出租方", f"要素填写错误：出租方法定代表人/负责人未填写完整")
    except:
        missObject += "要素“出租方法定代表人/负责人”缺失\n"

    try:
        match = '住所地：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&出租方住所地"] = factor
        if factor != "":
            factors_ok.append("主体&出租方住所地")
        else:
            factors_error["主体&出租方住所地"] = "住所地未填写完整"
            factors_to_inform["主体&出租方住所地"] = "出租方住所地未填写完整"
            addRemarkInDoc(word, document, "出租方", f"出租方住所地未填写完整")
    except:
        missObject += "要素“出租方住所地”缺失\n"

    try:
        match = '统一社会信用代码/身份证号码：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&出租方统一社会信用代码/身份证号码"] = factor
        if factor != "":
            if UnifiedSocialCreditIdentifier().check_code(factor, 'sc'):
                factors_ok.append("主体&出租方统一社会信用代码/身份证号码")
            else:
                # factors_error["主体&出租方统一社会信用代码/身份证号码"] = "主体&出租方统一社会信用代码/身份证号码未填写正确"
                if checkIdCard(factor) == 'ok':
                    factors_ok.append("主体&出租方统一社会信用代码/身份证号码")
                else:
                    rs = checkIdCard(factor)
                    factors_error["主体&出租方统一社会信用代码/身份证号码"] = "统一社会信用代码未填写正确或" + rs
                    factors_to_inform["主体&出租方统一社会信用代码/身份证号码"] = "出租方统一社会信用代码/身份证号码未填写完整"
                    addRemarkInDoc(word, document, "出租方", f"请核对并完善统一社会信用代码/身份证号码" + rs)
        else:
            factors_error["主体&出租方统一社会信用代码/身份证号码"] = "统一社会信用代码/身份证号码未填写完整"
            factors_to_inform["主体&出租方统一社会信用代码/身份证号码"] = "出租方统一社会信用代码/身份证号码未填写完整"
            addRemarkInDoc(word, document, "出租方", f"请核对并完善统一社会信用代码/身份证号码")
    except:
        missObject += "要素“出租方统一社会信用代码/身份证号码”缺失\n"

    try:
        match = '联系电话：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&出租方联系电话"] = factor
        if factor != "":
            if isTelPhoneNumber(factor) != "Error":
                factors_ok.append("主体&出租方联系电话")
            else:
                factors_error["主体&出租方联系电话"] = "联系电话填写有误"
                factors_to_inform["主体&出租方联系电话"] = "出租方联系电话未填写完整"
                addRemarkInDoc(word, document, "出租方", f"出租方联系电话填写错误")
        else:
            factors_error["主体&出租方联系电话"] = "联系电话未填写完整"
            factors_to_inform["主体&出租方联系电话"] = "出租方联系电话未填写完整"
            addRemarkInDoc(word, document, "出租方", f"出租方联系电话未填写完整")
    except:
        missObject += "要素“出租方联系电话”缺失\n"

    try:
        # 出租方电子邮箱检查
        match = '电子邮箱：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&出租方电子邮箱"] = factor
        if factor != "":
            if checkEmail(factor):
                factors_ok.append("主体&出租方电子邮箱")
            else:
                factors_error["主体&出租方电子邮箱"] = "出租方电子邮箱填写有误"
                factors_to_inform['主体&出租方电子邮箱审核提示'] = "请审核主体出租方电子邮箱是否填写正确"
                addRemarkInDoc(word, document, "出租方", f"出租方电子邮箱填写错误")
        else:
            factors_error["主体&出租方电子邮箱"] = "出租方电子邮箱未填写完整"
            factors_to_inform['主体&出租方电子邮箱审核提示'] = "请审核主体出租方电子邮箱是否填写正确"
            addRemarkInDoc(word, document, "出租方", f"出租方电子邮箱未填写完整")
    except:
        missObject += "要素“出租方电子邮箱”缺失\n"

    try:
        # 承租方主体审查
        match = '承租方：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&承租方"] = factor
        if factor != "":
            factors_ok.append("主体&承租方")
        else:
            factors_error["主体&承租方"] = "承租方未填写完整"
            factors_to_inform["主体&承租方"] = "承租方未填写完整"
            addRemarkInDoc(word, document, "承租方", f"承租方未填写完整")
    except:
        missObject += "要素“承租方”缺失\n"

    try:
        match = '法定代表人/负责人：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&承租方法定代表人/负责人"] = factor
        if factor != "":
            factors_ok.append("主体&承租方法定代表人/负责人")
        else:
            factors_error["主体&承租方法定代表人/负责人"] = "法定代表人/负责人未填写完整"
            factors_to_inform["主体&承租方法定代表人/负责人"] = "承租方法定代表人/负责人未填写完整"
            addRemarkInDoc(word, document, "承租方", f"承租方法定代表人/负责人未填写完整")
    except:
        missObject += "要素“承租方法定代表人/负责人”缺失\n"

    try:
        match = '住所地：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&承租方住所地"] = factor
        if factor != "":
            factors_ok.append("主体&承租方住所地")
        else:
            factors_error["主体&承租方住所地"] = "住所地未填写完整"
            factors_to_inform["主体&承租方住所地"] = "承租方住所地未填写完整"
            addRemarkInDoc(word, document, "承租方", f"承租方住所地未填写完整")
    except:
        missObject += "要素“承租方住所地”缺失\n"

    try:
        match = '统一社会信用代码/身份证号码：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&承租方统一社会信用代码/身份证号码"] = factor
        if factor != "":
            if UnifiedSocialCreditIdentifier().check_code(factor, 'sc'):
                factors_ok.append("主体&承租方统一社会信用代码/身份证号码")
            else:
                # factors_error["主体&出租方统一社会信用代码/身份证号码"] = "主体&出租方统一社会信用代码/身份证号码未填写正确"
                if checkIdCard(factor) == 'ok':
                    factors_ok.append("主体&承租方统一社会信用代码/身份证号码")
                else:
                    rs = checkIdCard(factor)
                    factors_error["主体&承租方统一社会信用代码/身份证号码"] = "统一社会信用代码未填写正确或" + rs
                    factors_to_inform["主体&承租方统一社会信用代码/身份证号码"] = "承租方统一社会信用代码/身份证号码未填写完整"
                    addRemarkInDoc(word, document, "承租方", f"请核对并完善统一社会信用代码/身份证号码或" + rs)

        else:
            factors_error["主体&承租方统一社会信用代码/身份证号码"] = "统一社会信用代码/身份证号码未填写完整"
            factors_to_inform["主体&承租方统一社会信用代码/身份证号码"] = "承租方统一社会信用代码/身份证号码未填写完整"
            addRemarkInDoc(word, document, "承租方", f"请核对并完善统一社会信用代码/身份证号码")
    except:
        missObject += "要素“承租方统一社会信用代码/身份证号码”缺失\n"

    try:
        match = '联系电话：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&承租方联系电话"] = factor
        if factor != "":
            if isTelPhoneNumber(factor) != "Error":
                factors_ok.append("主体&承租方联系电话")
            else:
                factors_error["主体&承租方联系电话"] = "联系电话填写有误"
                factors_to_inform["主体&承租方联系电话"] = "承租方联系电话未填写完整"
                addRemarkInDoc(word, document, "承租方", f"承租方联系电话填写有误")
        else:
            factors_error["主体&承租方联系电话"] = "联系电话未填写完整"
            factors_to_inform["主体&承租方联系电话"] = "出租方联系电话未填写完整"
            addRemarkInDoc(word, document, "承租方", f"承租方联系电话未填写完整")
    except:
        missObject += "要素“承租方联系电话”缺失\n"

    try:
        # 承租方电子邮箱检查
        match = '电子邮箱：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&承租方电子邮箱"] = factor
        if factor != "":
            if checkEmail(factor):
                factors_ok.append("主体&承租方电子邮箱")
            else:
                factors_error["主体&承租方电子邮箱"] = "承租方电子邮箱填写有误"
                factors_to_inform['主体&承租方电子邮箱审核提示'] = f"请审核主体承租方电子邮箱是否填写正确"
                addRemarkInDoc(word, document, "承租方", f"承租方邮箱填写错误")
        else:
            factors_error["主体&承租方电子邮箱"] = "承租方电子邮箱未填写完整"
            factors_to_inform['主体&电子邮箱审核提示'] = "请审核主体承租方电子邮箱是否填写正确"
            addRemarkInDoc(word, document, "承租方", f"承租方电子邮箱未填写完整")
    except:
        missObject += "要素“承租方电子邮箱”缺失\n"

    try:
        ##租赁房屋概况
        match = '1、房屋地址：【(.*?)】\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print("要素"+factor)
        factors["主体&房屋地址"] = factor
        if factor != "":
            factors_ok.append("主体&房屋地址")
        else:
            factors_error["主体&房屋地址"] = "房屋地址未填写完整"
            factors_to_inform['主体&房屋地址审核提示'] = "请审核主体房屋地址是否填写正确"
            addRemarkInDoc(word, document, "1、房屋地址", f"请审核主体房屋地址是否填写正确")
    except:
        missObject += "要素“房屋地址”缺失\n"

    try:
        ##租赁房屋概况、建筑面积
        match = '2、建筑面积：【(.*?)】'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print("要素"+factor)
        factors["主体&建筑面积"] = factor
        if factor != "":
            factors_ok.append("主体&建筑面积")
        else:
            factors_error["主体&建筑面积"] = "建筑面积未填写完整"
            factors_to_inform['主体&建筑面积审核提示'] = "请审核主体建筑面积是否填写正确"
            addRemarkInDoc(word, document, "2、建筑面积", f"提示：请审核主体建筑面积是否填写正确")
    except:
        missObject += "要素“建筑面积”缺失\n"

    try:
        ##租赁房屋概况、证书号码
        match = '证书号码：(.*?)】'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print("要素"+factor)
        factors["主体&证书号码"] = factor
        if factor != "":
            factors_ok.append("主体&证书号码")
        else:
            factors_error["主体&证书号码"] = "证书号码未填写完整"
            factors_to_inform['主体&证书号码审核提示'] = "请审核主体证书号码是否填写正确"
            addRemarkInDoc(word, document, "证书号码", f"提示：请审核主体证书号码是否填写正确")
    except:
        missObject += "要素“证书号码”缺失\n"

    try:
        ##租赁房屋概况、3.房屋用途
        match = '3、房屋用途：【(.*?)】'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print("要素"+factor)
        factors["主体&房屋用途"] = factor
        if factor != "":
            factors_ok.append("主体&房屋用途")
        else:
            factors_error["主体&房屋用途"] = "房屋用途未填写完整"
            factors_to_inform['主体&房屋用途审核提示'] = "请审核主体房屋用途是否填写正确"
            addRemarkInDoc(word, document, "3、房屋用途", f"提示：请审核主体房屋用途是否填写正确")
    except:
        missObject += "要素“房屋用途”缺失\n"

    try:
        ##租赁房屋概况、4、租赁房屋
        # .房屋用途
        match = '4、租赁房屋【(.*?)】（已/未二选一）设定抵押'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print("要素"+factor)
        factors["主体&4、租赁房屋"] = factor
        if factor != "":
            if factor == "已" or factor == "未":
                factors_ok.append("主体&租赁房屋")
            else:
                factors_error["主体&4、租赁房屋"] = "租赁房屋填写错误"
                factors_to_inform['主体&4、租赁房屋审核提示'] = "请审核主体租赁房屋是否填写正确"
                addRemarkInDoc(word, document, "4、租赁房屋", f"提示：请审核主体租赁房屋是否填写正确")
        else:
            factors_error["主体&4、租赁房屋"] = "租赁房屋未填写完整"
            factors_to_inform['主体&4、租赁房屋审核提示'] = "请审核主体租赁房屋是否填写完整"
            addRemarkInDoc(word, document, "4、租赁房屋", f"提示：请审核主体租赁房屋是否填写完整")
    except:
        missObject += "要素“4、租赁房屋”缺失\n"

    try:
        ##租赁房屋概况、5、租赁房屋
        # .房屋用途
        match = '5、租赁房屋【(.*?)】（已/未二选一）设定居住权'
        factor = re.findall(match, text)[0].replace(" ", "")
        # print("要素"+factor)
        factors["主体&5、租赁房屋"] = factor
        if factor != "":
            if factor == "已" or factor == "未":
                factors_ok.append("主体&5、租赁房屋")
            else:
                factors_error["主体&5、租赁房屋"] = "租赁房屋填写错误"
                factors_to_inform['主体&5、租赁房屋审核提示'] = "请审核主体租赁房屋是否填写正确"
                addRemarkInDoc(word, document, "5、租赁房屋", f"提示：请审核主体租赁房屋是否填写正确")
        else:
            factors_error["主体&5、租赁房屋"] = "租赁房屋未填写完整"
            factors_to_inform['主体&5、租赁房屋审核提示'] = "请审核主体租赁房屋是否填写完整"
            addRemarkInDoc(word, document, "5、租赁房屋", f"提示：请审核主体租赁房屋是否填写完整")
    except:
        missObject += "要素“5、租赁房屋”缺失\n"

    try:
        # 租赁期限
        match = '租赁期限共【(.*?)】年，从【(.*?)】年【(.*?)】月【(.*?)】日起至【(.*?)】年【(.*?)】月【(.*?)】日'
        factor = list(tuple(re.findall(match, text)[0]))
        date = list()
        for i in range(1, 4):
            date.append(factor[i])
        if len(factor) == 7:
            for i in range(1, 7):
                factor[i] = factor[i].replace(' ', '')
                factors["二&租赁期限check"] = f'{factor[1]}-{factor[2]}-{factor[3]}'
            factors["二&租赁期限"] = f'{factor[1]}-{factor[2]}-{factor[3]}起至{factor[4]}-{factor[5]}-{factor[6]}'
            if isRightDate(factor[1], factor[2], factor[3]) and isRightDate(factor[4], factor[5], factor[6]) and (
                    isRightDate(factor[1], factor[2], factor[3]) != isRightDate(factor[4], factor[5], factor[6])):
                factors_ok.append("二&租赁期限")
            else:
                factors_error["二&租赁期限"] = "请核对租赁期限是否准确"
                factors_to_inform['主体&租赁期限审核提示'] = "请审核主体租赁期限是否填写正确"
                addRemarkInDoc(word, document, "二、租赁期限", f"提示：请审核主体二&租赁期限是否填写正确")

        else:
            factors_error["二&租赁期限"] = "请核对租赁期限是否准确"
            factors_to_inform['主体&租赁期限审核提示'] = "请审核主体租赁期限是否填写正确"
            addRemarkInDoc(word, document, "二、租赁期限", f"提示：请审核主体二&租赁期限是否填写正确")
    except:
        missObject += "要素“二&租赁期限”缺失\n"

    try:
        # 三、租赁费用
        # 1、租金
        match = '金为人民币(.*?)元/(.*?)（大写：(.*?)元）。'
        factor = list(tuple(re.findall(match, text)[0]))
        factors["主体&三、租赁费用&1、租金"] = factor[0]
        if len(factor) == 3:
            if (factor[1] == '月' or factor[1] == '年' or factor[1] == '季度') and factor[2] == digital_to_Upper[factor[0]]:
                factors_ok.append("三、租赁费用&1、租金")
            else:
                factors_error["三、租赁费用&1、租金"] = "请财务部门确认租金具体数额是否准确"
                factors_to_inform['主体&三、租赁费用&1、租金审核提示'] = "请审核三、租赁费用&1、租金是否填写正确"
                addRemarkInDoc(word, document, "为人民币", f"提示：请审核三、租赁费用&1、租金是否填写正确")
        else:
            factors_error["三、租赁费用&1、租金"] = "请财务部门确认租金具体数额是否准确"
            factors_to_inform['主体&三、租赁费用&1、租金审核提示'] = "请审核三、租赁费用&1、租金是否填写正确"
            addRemarkInDoc(word, document, "为人民币", f"提示：请审核三、租赁费用&1、租金是否填写正确")
    except:
        missObject += "要素“三、租赁费用&1、租金”缺失\n"

    try:
        match = '租赁期限内，租金（是/否）【(.*?)】每年上调，上调幅度为：在上一个租期年度的基础上上浮【(.*?)】%。'
        factor2 = list(tuple(re.findall(match, text)[0]))
        factors["主体&三、租赁费用&1、租金是否上调"] = factor2[0]
        if (factor2[0] == '是' and factor2[1] != "") or factor2[0] == '否':
            factors_ok.append("三、租赁费用&1、租金")
        elif factor2[0] == '是' and factor2[1] == "":
            factors_error["三、租赁费用&1、租金"] = "请财务部门确认租金具体数额是否准确"
            factors_to_inform['主体&三、租赁费用&1、租金审核提示'] = "在“是”“否”中二选一"
            addRemarkInDoc(word, document, "基础上上浮", f"请检查租金上调幅度是否填写")
        else:
            factors_error["三、租赁费用&1、租金"] = "请财务部门确认租金具体数额是否准确"
            factors_to_inform['主体&三、租赁费用&1、租金审核提示'] = "在“是”“否”中二选一"
            addRemarkInDoc(word, document, "租金（是/否）", f"在“是”“否”中二选一")
    except:
        missObject += "要素“三、租赁费用&1、租金审核提示”缺失\n"

    try:
        match = '按一个月计算；租金按【(.*?)】（季度、半年、年三选一）缴纳'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&三、租金费用&2、租金计算方式"] = factor

        if factor == " ":
            factors_error["主体&三、租金费用&2、租金计算方式"] = "请检查租金计算方式是否填写完整"
            factors_to_inform["主体&三、租金费用&2、租金计算方式"] = "请检查租金计算方式是否填写完整"
            addRemarkInDoc(word, document, "按一个月计算；租金按", f"请检查租金计算方式是否填写完整")
        elif factor != "季度" or factor != "半年" or factor != "年":
            factors_ok.append("主体&三、租金费用&2、租金计算方式")
            factors_to_inform["主体&三、租金费用&2、租金计算方式"] = "请检查租金计算方式是否填写完整"
            addRemarkInDoc(word, document, "按一个月计算；租金按", f"请财务部门确定租金计算方式")
        # else:
        #     factors_error["主体&三、租金费用&2、租金计算方式"] = "请检查租金缴纳方式是否为季度、半年、年三选一"
        #     factors_to_inform["主体&三、租金费用&2、租金计算方式"] = "请检查租金计算方式是否填写完整"
        #     addRemarkInDoc(word, document, "按一个月计算；租金按", f"请财务部门确定租金计算方式")
    except:
        missObject += "要素“三、租金费用&2、租金计算方式”缺失\n"

    try:
        # 3、租金支付方式（1）
        match = '首期应在本合同签订后【(.*?)】日内付清，第二期开始，租金在前一租金支付期覆盖的租期届满之日前【(.*?)】日内一次性付清下一期租金。'
        factor = list(tuple(re.findall(match, text)[0]))
        if factor[0] != "" and factor[0] != "":
            factors_ok.append("三、租赁费用&3、租金支付方式（1）")
        else:
            factors_error["三、租赁费用&3、租金支付方式（1）"] = "请财务部门确认租金支付方式"
            factors_to_inform['三、租赁费用&2、租金计算方式（1）审核提示'] = "请财务部门确认租金支付方式"
            addRemarkInDoc(word, document, "首期应在本合同签订后", f"请财务部门确认租金支付方式")
    except:
        missObject += "要素“三、租赁费用&3、租金支付方式（1）”缺失\n"

    try:
        # 3、租金支付方式&（2）
        match = '开户行：【(.*?)】\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&开户行"] = factor
        if factor != "":
            factors_ok.append("主体&开户行")
        else:
            factors_error["主体&开户行"] = "请财务部门确认开户行是否准确"
            factors_to_inform['主体&开户行'] = "请财务部门确认开户行是否准确"
            addRemarkInDoc(word, document, "开户行", f"请财务部门确认开户行是否准确")
    except:
        missObject += "要素“开户行”缺失\n"

    try:
        # 3、租金支付方式&（2）
        match = '开户名：【(.*?)】\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&开户名"] = factor
        if factor != "":
            factors_ok.append("主体&开户名")
        else:
            factors_error["主体&开户名"] = "请财务部门确认开户名是否准确"
            factors_to_inform['主体&开户名'] = "请财务部门确认开户名是否准确"
            addRemarkInDoc(word, document, "开户名", f"请财务部门确认开户名是否准确")
    except:
        missObject += "要素“开户名”缺失\n"

    try:
        # 3、租金支付方式&（2）
        match = '银行账号：【(.*?)】\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&银行账号"] = factor
        if factor != "":
            factors_ok.append("主体&银行账号")
            factors_to_inform['主体&银行账号'] = "请财务部门确认银行账户是否准确"
            addRemarkInDoc(word, document, "银行账号", f"请财务部门确认银行账号是否准确")
        else:
            factors_error["主体&银行账号"] = "请财务部门确认银行账户是否准确"
            factors_to_inform['主体&银行账号'] = "请财务部门确认银行账户是否准确"
            addRemarkInDoc(word, document, "银行账号", f"请财务部门确认银行账号是否准确")
    except:
        missObject += "要素“银行账号”缺失\n"

    try:
        # 4、履约保证金
        match = '向出租方支付相当于【(.*?)】个月租金的款项'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&履约保证金"] = factor
        if factor in ("1", "2", "3"):
            factors_ok.append("主体&履约保证金")
        else:
            factors_error["主体&履约保证金"] = "请财务部门确认履约保证金是否准确"
            factors_to_inform['主体&履约保证金'] = "请财务部门确认履约保证金是否准确"
            addRemarkInDoc(word, document, "支付相当于", f"请财务部门确认履约保证金是否准确")
    except:
        missObject += "要素“履约保证金”缺失\n"

    try:
        # 4、租赁房屋交付时间
        match = '房屋交付时间定为【(.*?)】年【(.*?)】月【(.*?)】 日'
        factor = list(tuple(re.findall(match, text)[0]))
        date2 = list()
        for i in range(0, 3):
            date2.append(factor[i])

        if len(factor) == 3:
            for i in range(0, 3):
                factor[i] = factor[i].replace(' ', '')
            factors["四&租赁房屋交付时间"] = f'{factor[0]}-{factor[1]}-{factor[2]}'
            if isRightDate(factor[0], factor[1], factor[2]):
                # if date2 == date:
                #     factors_ok.append("四&租赁房屋交付时间")
                if factors["四&租赁房屋交付时间"] == factors["二&租赁期限check"]:
                    factors_ok.append("四&租赁房屋交付时间")
                else:
                    factors_error["四&租赁房屋交付时间"] = "请核对租赁房屋交付时间是否与第二条租赁期限起始日期一致"
                    factors_to_inform['主体&租赁房屋交付时间审核提示'] = "请审核主体租赁房屋交付时间是否与第二条租赁期限起始日期一致"
                    addRemarkInDoc(word, document, "房屋交付时间：", f"提示：请审核租赁房屋交付时间是否与第二条租赁期限起始日期一致")
            else:
                factors_error["四&租赁房屋交付时间"] = "请核对租赁房屋交付时间是否准确"
                factors_to_inform['主体&租赁房屋交付时间审核提示'] = "请审核主体租赁房屋交付时间是否填写完整"
                addRemarkInDoc(word, document, "房屋交付时间定为", f"提示：请审核租赁房屋交付时间是否填写完整")

        else:
            factors_error["四&租赁房屋交付时间"] = "请核对租赁房屋交付时间是否准确"
            factors_to_inform['主体&租赁房屋交付时间审核提示'] = "请审核主体租赁房屋交付时间是否填写正确"
            addRemarkInDoc(word, document, "房屋交付时间定为", f"提示：请审核主体四&房屋交付时间定为是否填写正确")
    except:
        missObject += "要素租赁房屋交付时间”缺失\n"

    try:
        match = '双方同意按以下第(.*?)种方式处'
        factor = re.findall(match, text)[0].replace(" ", "")
        print(factor)
        factors["主体&四&4"] = factor
        if factor != "":
            factors_ok.append("主体&四&4")
            if factor == "4" or "(4)" or "四":
                factors_to_inform["主体&四&4"] = "请核对第（4）项具体内容"
                addRemarkInDoc(word, document, "双方同意按以下第", f"是否填写房屋添附处理方式")
        else:
            factors_error["主体&四&4"] = "主体&四&4未填写完整"
            factors_to_inform["主体&四&4"] = "主体&四&4未填写完整"
            addRemarkInDoc(word, document, "双方同意按以下第", f"请检查是否填写房屋添附处理方式")
    except:
        missObject += "要素“主体&四&4 ”缺失\n"
        pass

    try:
        match = '甲方：(.*?)联系人'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&甲方"] = factor
        if factor != "":
            factors_ok.append("主体&甲方")
        else:
            factors_error["主体&甲方"] = "甲方未填写完整"
            factors_to_inform["主体&甲方"] = "甲方未填写完整"
            addRemarkInDoc(word, document, "甲方", f"要素填写错误：甲方未填写完整")
    except:
        missObject += "要素“甲方”缺失\n"

    try:
        match = '联系人：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&甲方联系人"] = factor
        if factor != "":
            factors_ok.append("主体&甲方联系人")
        else:
            factors_error["主体&甲方联系人"] = "甲方联系人未填写完整"
            factors_to_inform["主体&甲方联系人"] = "甲方联系人未填写完整"
            addRemarkInDoc(word, document, "甲方", f"要素填写错误：甲方联系人未填写完整")
    except:
        missObject += "要素“甲方联系人”缺失\n"

    try:
        match = '通信地址：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&甲方通信地址"] = factor
        if factor != "":
            factors_ok.append("主体&甲方通信地址")
        else:
            factors_error["主体&甲方通信地址"] = "甲方通信地址未填写完整"
            factors_to_inform["主体&甲方通信地址"] = "甲方通信地址未填写完整"
            addRemarkInDoc(word, document, "甲方", f"要素填写错误：甲方通信地址未填写完整")
    except:
        missObject += "要素“甲方通信地址”缺失\n"

    try:
        # 出租方电子邮箱检查
        match = '电子邮箱：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&甲方电子邮箱"] = factor
        if factor != "":
            if isEmail(factor):
                factors_ok.append("主体&甲方电子邮箱")
            else:
                factors_error["主体&甲方电子邮箱"] = "甲方电子邮箱填写有误"
                factors_to_inform['主体&甲方电子邮箱审核提示'] = "请审核主体甲方电子邮箱是否填写正确"
                addRemarkInDoc(word, document, "甲方", f"要素填写错误：甲方电子邮箱填写错误")
        else:
            factors_error["主体&甲方电子邮箱"] = "甲方电子邮箱未填写完整"
            factors_to_inform['主体&甲方电子邮箱审核提示'] = "请审核主体甲方电子邮箱是否填写正确"
            addRemarkInDoc(word, document, "甲方", f"要素填写错误：甲方电子邮箱未填写完整")
    except:
        missObject += "要素“甲方电子邮箱”缺失\n"

    try:
        match = '联系电话：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&甲方联系电话"] = factor
        if factor != "":
            if isTelPhoneNumber(factor) != "Error":
                factors_ok.append("主体&甲方联系电话")
            else:
                factors_error["主体&甲方联系电话"] = "甲方联系电话填写有误"
                factors_to_inform["主体&甲方联系电话"] = "甲方联系电话未填写完整"
                addRemarkInDoc(word, document, "甲方", f"要素填写错误：甲方联系电话填写有误")
        else:
            factors_error["主体&甲方联系电话"] = "甲方联系电话未填写完整"
            factors_to_inform["主体&甲方联系电话"] = "甲方联系电话未填写完整"
            addRemarkInDoc(word, document, "甲方", f"要素填写错误：甲方联系电话未填写完整")
    except:
        missObject += "要素“甲方联系电话”缺失\n"

    try:
        match = '乙方：(.*?)联系人'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&乙方"] = factor
        if factor != "":
            factors_ok.append("主体&乙方")
        else:
            factors_error["主体&乙方"] = "乙方未填写完整"
            factors_to_inform["主体&乙方"] = "乙方未填写完整"
            addRemarkInDoc(word, document, "乙方", f"要素填写错误：乙方未填写完整")
    except:
        missObject += "要素“乙方”缺失\n"

    try:
        match = '联系人：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&乙方联系人"] = factor
        if factor != "":
            factors_ok.append("主体&乙方联系人")
        else:
            factors_error["主体&乙方联系人"] = "乙方联系人未填写完整"
            factors_to_inform["主体&乙方联系人"] = "乙方联系人未填写完整"
            addRemarkInDoc(word, document, "乙方", f"要素填写错误：乙方联系人未填写完整")
    except:
        missObject += "要素“乙方联系人”缺失\n"

    try:
        match = '通信地址：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&乙方通信地址"] = factor
        if factor != "":
            factors_ok.append("主体&乙方通信地址")
        else:
            factors_error["主体&乙方通信地址"] = "乙方通信地址未填写完整"
            factors_to_inform["主体&乙方通信地址"] = "乙方通信地址未填写完整"
            addRemarkInDoc(word, document, "乙方", f"要素填写错误：乙方通信地址未填写完整")
    except:
        missObject += "要素“乙方通信地址”缺失\n"

    try:
        # 出租方电子邮箱检查
        match = '电子邮箱：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&乙方电子邮箱"] = factor
        if factor != "":
            if isEmail(factor):
                factors_ok.append("主体&乙方电子邮箱")
            else:
                factors_error["主体&乙方电子邮箱"] = "乙方电子邮箱填写有误"
                factors_to_inform['主体&乙方电子邮箱审核提示'] = "请审核主体乙方电子邮箱是否填写正确"
                addRemarkInDoc(word, document, "乙方", f"要素填写错误：乙方电子邮箱填写错误")
        else:
            factors_error["主体&乙方电子邮箱"] = "乙方电子邮箱未填写完整"
            factors_to_inform['主体&乙方电子邮箱审核提示'] = "请审核主体乙方电子邮箱是否填写正确"
            addRemarkInDoc(word, document, "乙方", f"要素填写错误：乙方电子邮箱未填写完整")
    except:
        missObject += "要素“乙方电子邮箱”缺失\n"

    try:
        match = '联系电话：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&乙方联系电话"] = factor
        if factor != "":
            if isTelPhoneNumber(factor) != "Error":
                factors_ok.append("主体&乙方联系电话")
            else:
                factors_error["主体&乙方联系电话"] = "乙方联系电话填写有误"
                factors_to_inform["主体&乙方联系电话"] = "乙方联系电话未填写完整"
                addRemarkInDoc(word, document, "乙方", f"要素填写错误：乙方联系电话填写有误")
        else:
            factors_error["主体&乙方联系电话"] = "乙方联系电话未填写完整"
            factors_to_inform["主体&乙方联系电话"] = "乙方联系电话未填写完整"
            addRemarkInDoc(word, document, "乙方", f"要素填写错误：乙方联系电话未填写完整")
    except:
        missObject += "要素“乙方联系电话”缺失\n"

    try:
        match = '双方同意向【(.*?)】'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&十一、合同的争议解决方法"] = factor
        if factor != "":
            if factor == factors["主体&房屋地址"]:

                factors_ok.append("主体&十一、合同的争议解决方法")
            else:
                factors_error["主体&十一、合同的争议解决方法"] = "主体&十一、合同的争议解决方法填写有误"
                factors_to_inform["主体&十一、合同的争议解决方法"] = "主体&十一、合同的争议解决方法未填写完整"
                addRemarkInDoc(word, document, "双方同意向", f"请检查主体&十一、合同的争议解决方法中法院是否位于房屋所在地")
        else:
            factors_error["主体&十一、合同的争议解决方法"] = "主体&十一、合同的争议解决方法未填写完整"
            factors_to_inform["主体&十一、合同的争议解决方法"] = "主体&十一、合同的争议解决方法未填写完整"
            addRemarkInDoc(word, document, "双方同意向", f"主体&十一、合同的争议解决方法未填写完整")
    except:
        missObject += "要素“主体&十一、合同的争议解决方法”缺失\n"

    try:
        # 租赁期限
        match = '一式【(.*?)】份，出租方【(.*?)】'
        factor = list(tuple(re.findall(match, text)[0]))
        number = list()
        nums = {0: '零', 1: '壹', 2: '贰', 3: '叁', 4: '肆', 5: '伍', 6: '陆', 7: '柒', 8: '捌', 9: '玖'}
        for i in range(0, 2):
            number.append(nums[factor[i]])
            # number[i].replace(" ", "")
        # if len(factor) == 2:

        if number[0] != "" and number[1] != "":
            if number[0] == (number[1] + 2):
                factors_ok.append("主体&十二、2、合同份数")
            else:
                factors_error["主体&十二、2、合同份数"] = "请核对主体&十二、2、合同份数是否准确"
                factors_to_inform['主体&十二、2、合同份数审核提示'] = "请审核主体&十二、2、合同份数是否填写正确"
                addRemarkInDoc(word, document, "一式", f"提示：请审核主体&十二、2、合同份数是否填写正确")
        else:
            factors_error["主体&十二、2、合同份数"] = "请核对主体&十二、2、合同份数是否准确"
            factors_to_inform['主体&十二、2、合同份数审核提示'] = "请审核主体&十二、2、合同份数是否填写正确"
            addRemarkInDoc(word, document, "一式", f"提示：请审核主体&十二、2、合同份数是否填写正确")
    except:
        missObject += "要素“主体&十二、2、合同份数”缺失\n"

    #
    try:
        match = '十三、补充条款：\n(.*?)出租方盖章'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&十三、补充条款"] = factor
        if factor != "":
            factors_ok.append("主体&十三、补充条款")
        else:
            factors_error["主体&十三、补充条款"] = "十三、补充条款未填写完整"
            factors_to_inform["主体&十三、补充条款"] = "十三、补充条款未填写完整"
            addRemarkInDoc(word, document, "十三、补充条款", f"请检查补充条款是否填写，若未有补充条款，应填写“无”")
    except:
        missObject += "要素“十三、补充条款”缺失\n"

    try:
        match = '法定代表人/授权代表：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&出租方法定代表人/授权代表"] = factor
        if factor != "":
            if factor == factors["主体&出租方法定代表人/负责人"]:
                factors_ok.append("主体&出租方法定代表人/授权代表")
            else:
                factors_error["主体&出租方法定代表人/授权代表"] = "法定代表人/授权代表与合同首部主体信息不一致"
                factors_to_inform["主体&出租方法定代表人/授权代表"] = "请核实法定代表人或负责人信息填写是否正确"
                addRemarkInDoc(word, document, "出租方(盖章)", f"请核实法定代表人或负责人信息填写是否正确")
        else:
            factors_error["主体&出租方法定代表人/授权代表"] = "法定代表人/授权代表未填写完整"
            factors_to_inform["主体&出租方法定代表人/授权代表"] = "出租方法定代表人/授权代表未填写完整"
            addRemarkInDoc(word, document, "出租方(盖章)", f"出租方法定代表人/授权代表未填写完整")
    except:
        missObject += "要素“出租方法定代表人/授权代表”缺失\n"

    # try:
    #     #出租方盖章日期
    #     match = '(.*?)年(.*?)月(.*?)日'
    #     factor = list(tuple(re.findall(match, text)[0]))
    #
    #     if len(factor) == 3:
    #         for i in range(0, 3):
    #             factor[i] = factor[i].replace(' ', '')
    #         factors["出租方盖章日期"] = f'{factor[0]}-{factor[1]}-{factor[2]}'
    #         if isRightDate(factor[0], factor[1], factor[2]):
    #
    #                 factors_ok.append("出租方盖章日期")
    #         else:
    #             factors_error["出租方盖章日期"] = "请核对出租方盖章日期是否准确"
    #             factors_to_inform['主体&出租方盖章日期审核提示'] = "请审核出租方盖章日期是否填写正确"
    #             addRemarkInDoc(word, document, "出租方(盖章)", f"出租方盖章日期是否填写正确")
    #
    #     else:
    #         factors_error["出租方盖章日期"] = "请核对租赁房屋交付时间是否准确"
    #         factors_to_inform['主体&出租方盖章日期审核提示'] = "请审核主体租赁房屋交付时间是否填写正确"
    #         addRemarkInDoc(word, document, "出租方(盖章)", f"出租方盖章日期是否填写正确")
    # except:
    #     missObject += "要素出租方盖章日期是否填写正确”缺失\n"

    try:
        match = '法定代表人/授权代表：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&承租方法定代表人/授权代表"] = factor
        if factor != "":
            if factor == factors["主体&承租方法定代表人/负责人"]:
                factors_ok.append("主体&承租方法定代表人/授权代表")
            else:
                factors_error["主体&承租方法定代表人/授权代表"] = "法定代表人/授权代表与合同首部主体信息不一致"
                factors_to_inform["主体&承租方法定代表人/授权代表"] = "请核实法定代表人或负责人信息填写是否正确"
                addRemarkInDoc(word, document, "承租方(盖章)", f"请核实法定代表人或负责人信息填写是否正确")
        else:
            factors_error["主体&承租方法定代表人/授权代表"] = "法定代表人/授权代表未填写完整"
            factors_to_inform["主体&承租方法定代表人/授权代表"] = "承租方法定代表人/授权代表未填写完整"
            addRemarkInDoc(word, document, "承租方(盖章)", f"承租方法定代表人/授权代表未填写完整")
    except:
        missObject += "要素“承租方法定代表人/授权代表”缺失\n"

    # try:
    #     # 承租方盖章日期
    #     match = '(.*?)年(.*?)月(.*?)日附件'
    #     factor = list(tuple(re.findall(match, text)[0]))
    #
    #     if len(factor) == 3:
    #         for i in range(0, 3):
    #             factor[i] = factor[i].replace(' ', '')
    #         factors["承租方盖章日期"] = f'{factor[0]}-{factor[1]}-{factor[2]}'
    #         if isRightDate(factor[0], factor[1], factor[2]):
    #
    #             factors_ok.append("承租方盖章日期")
    #         else:
    #             factors_error["承租方盖章日期"] = "请核对承租方盖章日期是否准确"
    #             factors_to_inform['主体&承租方盖章日期审核提示'] = "请审核承租方盖章日期是否填写正确"
    #             addRemarkInDoc(word, document, "承租方(盖章) ", f"承租方盖章日期是否填写正确")
    #
    #     else:
    #         factors_error["承租方盖章日期"] = "请核对租赁房屋交付时间是否准确"
    #         factors_to_inform['主体&承租方盖章日期审核提示'] = "请审核主体租赁房屋交付时间是否填写正确"
    #         addRemarkInDoc(word, document, "承租方(盖章) ", f"承租方盖章日期是否填写正确")
    # except:
    #     missObject += "要素承租方盖章日期是否填写正确”缺失\n"

    try:
        if missObject != "":
            addRemarkInDoc(word, document, "", missObject)
        copy_path = processed_file_sava_dir + "/" + filePath.split("/")[-1]
        filePath = str_insert(copy_path, copy_path.index(".doc"), "(已审查)")
        print(filePath)
        document.SaveAs(filePath)
        document.Close()
        factors1, factors_ok1, factors_error1, factors_to_inform1, word = None_standard_contract.lease_contract(
            filePath,
            processed_file_sava_dir)
        os.remove(filePath)
        # word.Quit()
    except Exception as ex:
        print(ex)
    print("中间文件已删除")
    print(factors, factors_ok, factors_error, factors_to_inform)

    return factors, factors_ok, factors_error, factors_to_inform
