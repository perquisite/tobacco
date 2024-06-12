import os
import re

import pythoncom
from win32com.client import Dispatch
import None_standard_contract
from utils import UnifiedSocialCreditIdentifier, checkIdCard, isTelPhoneNumber, isRightDate, checkQQ, checkEmail, \
    str_insert, addRemarkInDoc, digital_to_Upper, is_contain_dot, isEmail, checkEntersAndSpace

import helpful as hp


# add by qy
def processFunc(text, tables, filePath, processed_file_sava_dir):
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
        document = word.Documents.Open(FileName=filePath,)
    except Exception as ex:
        print(ex)
    factors = {}
    factors_ok = []
    factors_error = {}
    factors_to_inform = {}
    # print(text, tables)

    # 缺失的要素
    missObject = ""

    # 甲方主体审查
    try:
        match = '甲方：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&甲方"] = factor
        if factor != "":
            factors_ok.append("主体&甲方")
        else:
            factors_error["主体&甲方"] = "甲方未填写完整"
            addRemarkInDoc(word, document, "甲方", "要素填写错误：甲方未填写完整")
    except:
        missObject += "要素“甲方”缺失\n"

    try:
        match = '甲方法定代表人/负责人：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&甲方法定代表人/负责人"] = factor
        if factor != "":
            factors_ok.append("主体&甲方法定代表人/负责人")
        else:
            factors_error["主体&甲方法定代表人/负责人"] = "法定代表人/负责人未填写完整"
            addRemarkInDoc(word, document, "甲方法定代表人/负责人", "要素填写错误：甲方法定代表人/负责人未填写完整")
    except:
        missObject += "要素“甲方法定代表人/负责人”缺失\n"

    try:
        match = '甲方住所地：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&甲方住所地"] = factor
        if factor != "":
            factors_ok.append("主体&甲方住所地")
        else:
            factors_error["主体&甲方住所地"] = "住所地未填写完整"
            addRemarkInDoc(word, document, "甲方住所地", "要素填写错误：住所地未填写完整")
    except:
        missObject += "要素“甲方住所地”缺失\n"

    try:
        match = '甲方统一社会信用代码/身份证号码：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&甲方统一社会信用代码/身份证号码"] = factor
        if factor != "":
            if UnifiedSocialCreditIdentifier().check_code(factor, 'sc'):
                factors_ok.append("主体&甲方统一社会信用代码/身份证号码")
            else:
                # factors_error["主体&甲方统一社会信用代码/身份证号码"] = "主体&甲方统一社会信用代码/身份证号码未填写正确"
                if checkIdCard(factor) == 'ok':
                    factors_ok.append("主体&甲方统一社会信用代码/身份证号码")
                else:
                    rs = checkIdCard(factor)
                    factors_error["主体&甲方统一社会信用代码/身份证号码"] = "统一社会信用代码未填写正确或" + rs
                    addRemarkInDoc(word, document, "甲方统一社会信用代码/身份证号码", "要素填写错误：统一社会信用代码未填写正确或" + rs)
        else:
            factors_error["主体&甲方统一社会信用代码/身份证号码"] = "统一社会信用代码/身份证号码未填写完整"
            addRemarkInDoc(word, document, "甲方统一社会信用代码/身份证号码", "要素填写错误：统一社会信用代码/身份证号码未填写完整")
    except:
        missObject += "要素“甲方统一社会信用代码/身份证号码”缺失\n"

    try:
        match = '甲方联系方式：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&甲方联系方式"] = factor
        if factor != "":
            if isTelPhoneNumber(factor)!="Error":
                factors_ok.append("主体&甲方联系方式")
            else:
                factors_error["主体&甲方联系方式"] = "联系方式填写有误"
                addRemarkInDoc(word, document, "甲方联系方式", "要素填写错误：联系方式填写有误")
        else:
            factors_error["主体&甲方联系方式"] = "联系方式未填写完整"
            addRemarkInDoc(word, document, "甲方联系方式", "要素填写错误：联系方式未填写完整")
    except:
        missObject += "要素“甲方联系方式”缺失\n"

    try:
        # 乙方主体审查
        match = '乙方：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&乙方"] = factor
        if factor != "":
            factors_ok.append("主体&乙方")
        else:
            factors_error["主体&乙方"] = "乙方未填写完整"
            addRemarkInDoc(word, document, "乙方", "要素填写错误：乙方未填写完整")
    except:
        missObject += "要素“乙方”缺失\n"

    try:
        match = '乙方法定代表人/负责人：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&乙方法定代表人/负责人"] = factor
        if factor != "":
            factors_ok.append("主体&乙方法定代表人/负责人")
        else:
            factors_error["主体&乙方法定代表人/负责人"] = "法定代表人/负责人未填写完整"
            addRemarkInDoc(word, document, "乙方法定代表人/负责人", "要素填写错误：法定代表人/负责人未填写完整")
    except:
        missObject += "要素“乙方法定代表人/负责人”缺失\n"

    try:
        match = '乙方住所地：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&乙方住所地"] = factor
        if factor != "":
            factors_ok.append("主体&乙方住所地")
        else:
            factors_error["主体&乙方住所地"] = "住所地未填写完整"
            addRemarkInDoc(word, document, "乙方住所地", "要素填写错误：住所地未填写完整")
    except:
        missObject += "要素“乙方住所地”缺失\n"

    try:
        match = '乙方统一社会信用代码/身份证号码：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&乙方统一社会信用代码/身份证号码"] = factor
        if factor != "":
            if UnifiedSocialCreditIdentifier().check_code(factor, 'sc'):
                factors_ok.append("主体&乙方统一社会信用代码/身份证号码")
            else:
                # factors_error["主体&甲方统一社会信用代码/身份证号码"] = "主体&甲方统一社会信用代码/身份证号码未填写正确"
                if checkIdCard(factor) == 'ok':
                    factors_ok.append("主体&乙方统一社会信用代码/身份证号码")
                else:
                    rs = checkIdCard(factor)
                    factors_error["主体&乙方统一社会信用代码/身份证号码"] = "统一社会信用代码未填写正确或" + rs
                    addRemarkInDoc(word, document, "乙方统一社会信用代码/身份证号码", "要素填写错误：统一社会信用代码未填写正确或" + rs)

        else:
            factors_error["主体&乙方统一社会信用代码/身份证号码"] = "统一社会信用代码/身份证号码未填写完整"
            addRemarkInDoc(word, document, "乙方统一社会信用代码/身份证号码", "要素填写错误：统一社会信用代码/身份证号码未填写完整")
    except:
        missObject += "要素“乙方统一社会信用代码/身份证号码”缺失\n"

    try:
        match = '乙方联系方式：(.*?)\n'
        factor = re.findall(match, text)[0].replace(" ", "")
        factors["主体&乙方联系方式"] = factor
        if factor != "":
            if isTelPhoneNumber(factor)!="Error":
                factors_ok.append("主体&乙方联系方式")
            else:
                factors_error["主体&乙方联系方式"] = "联系方式填写有误"
                addRemarkInDoc(word, document, "乙方联系方式", "要素填写错误：联系方式填写有误")
        else:
            factors_error["主体&乙方联系方式"] = "联系方式未填写完整"
            addRemarkInDoc(word, document, "乙方联系方式", "要素填写错误：联系方式未填写完整")
    except:
        missObject += "要素“乙方联系方式”缺失\n"

    try:
        # （招标编号：xxx）的xxx供应商入库事项
        match = '（招标编号：(.*?)）的(.*?)供'
        factor = list(tuple(re.findall(match, text)[0]))
        if len(factor) == 2:
            factor[0] = factor[0].replace(' ', '')
            factor[1] = factor[1].replace(' ', '')
            factors["主体&招标编号"] = factor[0]
            if factor[0] != "":
                factors_ok.append("主体&招标编号")
            else:
                factors_error["主体&招标编号"] = "招标编号未填写完整"
                addRemarkInDoc(word, document, "（招标编号：", "要素填写错误：招标编号未填写完整")

            if factor[1] != "":
                factors_ok.append("主体&招标项目")
            else:
                factors_error["主体&招标项目"] = "招标项目未填写完整"
                addRemarkInDoc(word, document, "供应商入库事项协商一致", "要素填写错误：招标项目未填写完整")
        else:
            factors_error["主体&招标编号和项目"] = "（招标编号：xxx）的xxx供应商入库事项未填写完整"
            addRemarkInDoc(word, document, "招标编号", "要素填写错误：招标编号和招标项目未填写完整")

        factors_to_inform['主体&主体审核提示'] = "请审核主体相关信息是否正确"
        addRemarkInDoc(word, document, "协商一致，同意按下述条款和条件签署本合同", f"提示：请审核主体相关信息是否正确")
    except:
        missObject += "要素“（招标编号：xxx）的xxx供应商入库事项”缺失\n"

    try:
        # 协议期限
        match = '协议期限：从(.*?)年(.*?)月(.*?)日至(.*?)年(.*?)月(.*?)日'
        factor = list(tuple(re.findall(match, text)[0]))
        if len(factor) == 6:
            for i in range(6):
                factor[i] = factor[i].replace(' ', '')
            factors["三1&协议期限"] = f'{factor[0]}-{factor[1]}-{factor[2]}至{factor[3]}-{factor[4]}-{factor[5]}'
            if isRightDate(factor[0], factor[1], factor[2]) and isRightDate(factor[3], factor[4], factor[5]):
                factors_ok.append("三1&协议期限")
            else:
                factors_error["三1&协议期限"] = "协议期限时间未填写规范"
                addRemarkInDoc(word, document, "协议期限", "要素填写错误：协议期限时间未填写规范")
        else:
            factors_error["三1&协议期限"] = "协议期限时间未填写完整"
            addRemarkInDoc(word, document, "协议期限", "要素填写错误：协议期限时间未填写完整")
        # 处理表格的内容,至少保证一条记录
        if len(tables) == 1:
            r_num = 0
            # 审查第一个表，”乙方为甲方提供固定的项目负责人及具体服务人员“
            table_1 = tables[0]
            for i in range(0, len(table_1.rows)):
                if i > 1:
                    # 每一行一个要素，三条第3小点表格第三行
                    flgStr = []
                    passFlag = True
                    name = table_1.cell(i, 0).text.replace(' ', '')
                    if name == "":
                        passFlag = False
                        flgStr.append('姓名填写不完整')

                    idCard = table_1.cell(i, 1).text.replace(' ', '')
                    if idCard == "":
                        passFlag = False
                        flgStr.append('证件号码填写不完整')
                    else:
                        if checkIdCard(idCard) != 'ok':
                            flgStr.append('证件号码：' + checkIdCard(idCard))
                            passFlag = False

                    role = table_1.cell(i, 2).text.replace(' ', '')
                    if role == "":
                        passFlag = False
                        flgStr.append('乙方单位职务填写不完整')

                    role1 = table_1.cell(i, 3).text.replace(' ', '')
                    if role1 == "":
                        passFlag = False
                        flgStr.append('本项目承担职务填写不完整')

                    phone = table_1.cell(i, 4).text.replace(' ', '')
                    if phone == "":
                        passFlag = False
                        flgStr.append('电话填写不完整')
                    else:
                        if isTelPhoneNumber(phone) == 'Error':
                            flgStr.append('电话号码格式不规范')
                            passFlag = False

                    qq = table_1.cell(i, 5).text.replace(' ', '')
                    if qq == "":
                        passFlag = False
                        flgStr.append('QQ填写不完整')
                    else:
                        if checkQQ(qq) == False:
                            flgStr.append('QQ格式不规范')
                            passFlag = False
                    email = table_1.cell(i, 6).text.replace(' ', '')
                    if email == "":
                        passFlag = False
                        flgStr.append('邮箱填写不完整')
                    else:
                        if checkEmail(email) == False:
                            flgStr.append('邮箱格式不规范')
                            passFlag = False

                    other = table_1.cell(i, 7).text.replace(' ', '')
                    ##暂时不处理

                    factors["三3行" + str(
                        i + 1) + "&项目负责人及具体服务人员表"] = f"{name}、{idCard}、{role}、{role1}、{phone}、{qq}、{email}、{other}"
                    # 先判断一下该行是不是空的
                    if name or idCard or role or role1 or phone or qq or email or other:
                        r_num += 1
                    else:
                        continue
                    if passFlag:
                        factors_ok.append("三3行" + str(i - 4) + "&项目负责人及具体服务人员表")
                    else:
                        err_str = ''
                        for j in range(0, len(flgStr)):
                            if j == 0:
                                err_str += flgStr[j]
                            else:
                                err_str += '、' + flgStr[j]
                        factors_error["三3行" + str(i - 4) + "&项目负责人及具体服务人员表"] = err_str

                        addRemarkInDoc(word, document, "乙方为甲方提供固定的项目负责人及具体服务人员", f"要素填写错误：表格第{str(i)}行{err_str}")

            if r_num == 0:
                factors_error["三3&项目负责人及具体服务人员表"] = "请补充项目负责人及具体服务人员表至少一项"
                addRemarkInDoc(word, document, "乙方为甲方提供固定的项目负责人及具体服务人员", f"要素填写错误：请补充项目负责人及具体服务人员表至少一项")

        else:
            factors_error["三3&项目负责人及具体服务人员表"] = "缺少项目负责人及具体服务人员表"
            addRemarkInDoc(word, document, "乙方为甲方提供固定的项目负责人及具体服务人员", f"要素填写错误：缺少项目负责人及具体服务人员表")

        factors_to_inform['三3&表格审核提示'] = "请审核表格相关填写信息是否正确"
        addRemarkInDoc(word, document, "乙方为甲方提供固定的项目负责人及具体服务人员", f"提示：请审核表格相关填写信息是否正确")
    except:
        missObject += "要素“乙方为甲方提供固定的项目负责人及具体服务人员乙方为甲方提供固定的项目负责人及具体服务人员”表格缺失\n"

    try:
        # 甲方对乙方相关管理依据
        match = '按照甲方对供应商库的相关管理要求对乙方进行管理具体文件如下：([\s\S]*?)五'
        factor = re.findall(match, text)[0].replace(" ", "")
        print("要素"+factor)
        factors["四3&甲方对乙方相关管理依据"] = factor
        if checkEntersAndSpace(factor):
            factors_ok.append("四3&甲方对乙方相关管理依据")
            factors_to_inform['四3&甲方对乙方相关管理依据'] = "请审核管理依据是否正确、完整"
            addRemarkInDoc(word, document, "按照甲方对供应商库的相关管理", f"提示：请核实该项约定的必要性和合理性")
        else:
            factors_error["四3&甲方对乙方相关管理依据"] = "按照甲方对供应商库的相关管理要求未填写完整"
            addRemarkInDoc(word, document, "按照甲方对供应商库的相关管理", f"要素填写错误：按照甲方对供应商库的相关管理要求未填写完整")
    except Exception as ex:
        print(ex)
        missObject += "要素“按照甲方对供应商库的相关管理要求对乙方进行管理具体文件如下”缺失\n"

    try:
        # 供货/服务范围及质量其他要求
        match = '3、其他(.*?)\n'
        # print(re.findall(match, text))
        # factor = re.findall(match, text)[0].replace(" ", "")
        factors["七3&供货/服务范围及质量其他要求"] = factor
        if factor != "":
            factors_to_inform['四3&供货/服务范围及质量其他要求'] = "请核实该项约定的必要性和合理性"
            addRemarkInDoc(word, document, "3、其他", f"提示：请核实该项约定的必要性和合理性")
    except Exception as ex:
        print(ex)
        missObject += "要素“3、其他”缺失\n"

    try:
        if missObject != "":
            addRemarkInDoc(word, document,"", missObject)
        copy_path = processed_file_sava_dir + "/" + filePath.split("/")[-1]
        filePath = str_insert(copy_path, copy_path.index(".doc"), "(已审查)")
        print(filePath)
        document.SaveAs(filePath)
        document.Close()
        factors1, factors_ok1, factors_error1, factors_to_inform1, word = None_standard_contract.purchase_contract(
            filePath,
            processed_file_sava_dir)
        os.remove(filePath)
        # word.Quit()
        print("中间文件已删除")
    except Exception as ex:
        print(ex)
    # print(factors, factors_ok, factors_error, factors_to_inform)
    # print('不是我卡的~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~标准-1')
    return factors, factors_ok, factors_error, factors_to_inform
