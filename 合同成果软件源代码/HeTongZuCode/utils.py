# -*- coding:utf-8 -*-
# @ModuleName: utils
# @Function:
# @Author: qinyang
# @email: 835343249@qq.com
# @Time: 2021/7/7 10:32
import datetime
import re
import rmbTrans

# 统一社会信用代码 + 组织结构代码校验
class UnifiedSocialCreditIdentifier(object):
    '''
    统一社会信用代码 + 组织结构代码校验
    '''

    def __init__(self):
        '''
        Constructor
        '''
        # 统一社会信用代码中不使用I,O,S,V,Z
        # ''.join([str(i) for i in range(10)])
        # import string
        # string.ascii_uppercase  # ascii_lowercase |  ascii_letters
        # dict([i for i in zip(list(self.string), range(len(self.string)))])
        # dict(enumerate(self.string))
        # list(d.keys())[list(d.values()).index(10)]
        # chr(97)  --> 'a'
        self.string1 = '0123456789ABCDEFGHJKLMNPQRTUWXY'
        self.SOCIAL_CREDIT_CHECK_CODE_DICT = {
            '0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9,
            'A': 10, 'B': 11, 'C': 12, 'D': 13, 'E': 14, 'F': 15, 'G': 16, 'H': 17,
            'J': 18, 'K': 19, 'L': 20, 'M': 21, 'N': 22, 'P': 23, 'Q': 24,
            'R': 25, 'T': 26, 'U': 27, 'W': 28, 'X': 29, 'Y': 30}
        # 第i位置上的加权因子
        self.social_credit_weighting_factor = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]

        # GB11714-1997全国组织机构代码编制规则中代码字符集
        self.string2 = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        self.ORGANIZATION_CHECK_CODE_DICT = {
            '0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9,
            'A': 10, 'B': 11, 'C': 12, 'D': 13, 'E': 14, 'F': 15, 'G': 16, 'H': 17, 'I': 18,
            'J': 19, 'K': 20, 'L': 21, 'M': 22, 'N': 23, 'O': 24, 'P': 25, 'Q': 26,
            'R': 27, 'S': 28, 'T': 29, 'U': 30, 'V': 31, 'W': 32, 'X': 33, 'Y': 34, 'Z': 35}
        # 第i位置上的加权因子
        self.organization_weighting_factor = [3, 7, 9, 10, 5, 8, 4, 2]

    # 统一社会信用代码 sc
    def check_social_credit_code(self, code):
        '''
        统一社会信用代码校验
        国家标准GB32100—2015：18位统一社会信用代码从2015年10月1日正式实行，
        标准规定统一社会信用代码用18位阿拉伯数字或大写英文字母（不使用I、O、Z、S、V）表示，
        分别是1位登记管理部门代码、1位机构类别代码、6位登记管理机关行政区划码、9位主体标识码（组织机构代码）、1位校验码


        税号 = 6位行政区划码 + 9位组织机构代码
        计算校验码公式:
            C18 = 31-mod(sum(Ci*Wi)，31)
        其中Ci为组织机构代码的第i位字符,Wi为第i位置的加权因子,C18为校验码
        c18=30, Y; c18=31, 0
        '''
        # 主要是避免缺失值乱入
        # if type(code) != str: return False
        # 转大写
        code = code.upper()
        # 1. 长度限制
        if len(code) != 18:
            #print('{} -- 统一社会信用代码长度不等18！'.format(code))
            return False
        # 2. 不含IOSVZ -- 组成限制, 非字典表给个非常大的数, 不超过15000
        '''lst = list('IOSVZ')
        for s in lst:
            if s in code:
                print('包含非组成字符：%s' % (s))
                return False'''

        # 2. 组成限制
        # 登记管理部门：1=机构编制; 5=民政; 9=工商; Y=其他
        # 机构类别代码:
        '''
        机构编制=1：1=机关 | 2=事业单位 | 3=中央编办直接管理机构编制的群众团体 | 9=其他
        民政=5：1=社会团体 | 2=民办非企业单位 | 3=基金会 | 9=其他
        工商=9：1=企业 | 2=个体工商户 | 3=农民专业合作社
        其他=Y：1=其他
        '''
        reg = r'^(11|12|13|19|51|52|53|59|91|92|93|Y1)\d{6}\w{9}\w$'
        if not re.match(reg, code):
            # print('{} -- 组成错误！'.format(code))
            return False

        # 3. 校验码验证
        # 本体代码
        ontology_code = code[:17]
        # 校验码
        check_code = code[17]
        # 计算校验码
        tmp_check_code = self.gen_check_code(self.social_credit_weighting_factor,
                                             ontology_code,
                                             31,
                                             self.SOCIAL_CREDIT_CHECK_CODE_DICT)
        if tmp_check_code == -1:
            #print('{} -- 包含非组成字符！'.format(code))
            return False

        tmp_check_code = (0 if tmp_check_code == 31 else tmp_check_code)
        if self.string1[tmp_check_code] == check_code:
            # print('{} -- 统一社会信用代码校验正确！'.format(code))
            return True
        else:
            #print('{} -- 统一社会信用代码校验错误！'.format(code))
            return False

    # 组织结构代码校验 org
    def check_organization_code(self, code):
        '''
        组织机构代码校验
        该规则按照GB 11714编制：统一社会信用代码的第9~17位为主体标识码(组织机构代码)，共九位字符
        计算校验码公式:
            C9 = 11-mod(sum(Ci*Wi)，11)
        其中Ci为组织机构代码的第i位字符,Wi为第i位置的加权因子,C9为校验码
        C9=10, X; C9=11, 0
        @param  code: 统一社会信用代码 / 组织机构代码
        '''
        # 主要是避免缺失值乱入
        # if type(code) != str: return False
        # 1. 长度限制
        if len(code) != 9:
            print('{} -- 组织机构代码长度不等9！'.format(code))
            return False

        # 2. 组成限制
        reg = r'^\w{9}$'
        if not re.match(reg, code):
            print('{} -- 组成错误！'.format(code))
            return False

        # 3. 校验码验证
        # 本体代码
        ontology_code = code[:8]
        # 校验码
        check_code = code[8]
        # 计算校验码
        tmp_check_code = self.gen_check_code(self.organization_weighting_factor,
                                             ontology_code,
                                             11,
                                             self.ORGANIZATION_CHECK_CODE_DICT)
        if tmp_check_code == -1:
            print('{} -- 包含非组成字符！'.format(code))
            return False

        tmp_check_code = (0 if tmp_check_code == 11
                          else (33 if tmp_check_code == 10 else tmp_check_code))
        if self.string2[tmp_check_code] == check_code:
            # print('{} -- 组织机构代码校验正确！'.format(code))
            return True
        else:
            print('{} -- 组织机构代码校验错误！'.format(code))
            return False

    def check_code(self, code, code_type='sc'):
        '''Series类型
        @code_type {org, sc}'''
        # try:
        if type(code) != str: return False
        if code_type == 'sc':
            return self.check_social_credit_code(code)
        elif code_type == 'org':
            return self.check_organization_code(code)
        else:
            if len(code) == 18:
                return self.check_social_credit_code(code)
            else:
                return self.check_organization_code(code) if len(code) == 9 else False
        # except Exception as err:
        #    print(err)
        #    print('code:', code)

    def gen_check_code(self, weighting_factor, ontology_code, modulus, check_code_dict):
        '''
        @param weighting_factor: 加权因子
        @param ontology_code:本体代码
        @param modulus:  模数(求余用)
        @param check_code_dict: 字符字典
        '''
        total = 0
        for i in range(len(ontology_code)):
            if ontology_code[i].isdigit():
                # print(ontology_code[i], weighting_factor[i])
                total += int(ontology_code[i]) * weighting_factor[i]
            else:
                num = check_code_dict.get(ontology_code[i], -1)
                if num < 0: return -1
                total += num * weighting_factor[i]
        diff = modulus - total % modulus
        # print(diff)
        return diff


# 身份证号校验 Errors= ['ok', '身份证号码位数不对!', '身份证号码出生日期超出范围或含有非法字符!', '身份证号码校验错误!', '身份证号码地区非法!']
def checkIdCard(id_code):
    Errors = ['ok', '身份证号码位数不对', '身份证号码出生日期超出范围或含有非法字符', '身份证号码校验错误', '身份证号码地区非法']
    area = {"11": "北京", "12": "天津", "13": "河北", "14": "山西", "15": "内蒙古", "21": "辽宁", "22": "吉林", "23": "黑龙江",
            "31": "上海",
            "32": "江苏", "33": "浙江", "34": "安徽", "35": "福建", "36": "江西", "37": "山东", "41": "河南", "42": "湖北", "43": "湖南",
            "44": "广东", "45": "广西", "46": "海南", "50": "重庆", "51": "四川", "52": "贵州", "53": "云南", "54": "西藏", "61": "陕西",
            "62": "甘肃", "63": "青海", "64": "宁夏", "65": "新疆", "71": "台湾", "81": "香港", "82": "澳门", "91": "国外"}
    id_code = str(id_code)
    id_code = id_code.strip()
    id_code_list = list(id_code)

    # 地区校验
    key = id_code[0: 2]  # TODO： cc  地区中的键是否存在
    if key in area.keys():
        if (not area[(id_code)[0:2]]):
            return Errors[4]
    else:
        return Errors[4]
    # 15位身份号码检测

    if (len(id_code) == 15):
        if ((int(id_code[6:8]) + 1900) % 4 == 0 or (
                (int(id_code[6:8]) + 1900) % 100 == 0 and (int(id_code[6:8]) + 1900) % 4 == 0)):
            erg = re.compile(
                '[1-9][0-9]{5}[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))[0-9]{3}$')  # //测试出生日期的合法性
        else:
            ereg = re.compile(
                '[1-9][0-9]{5}[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|1[0-9]|2[0-8]))[0-9]{3}$')  # //测试出生日期的合法性
        if (re.match(ereg, id_code)):
            return Errors[0]
        else:
            return Errors[2]
    # 18位身份号码检测
    elif (len(id_code) == 18):
        # 出生日期的合法性检查
        # 闰年月日:((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))
        # 平年月日:((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|1[0-9]|2[0-8]))
        if (int(id_code[6:10]) % 4 == 0 or (int(id_code[6:10]) % 100 == 0 and int(id_code[6:10]) % 4 == 0)):
            ereg = re.compile(
                '[1-9][0-9]{5}(19[0-9]{2}|20[0-9]{2})((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))[0-9]{3}[0-9Xx]$')  # //闰年出生日期的合法性正则表达式
        else:
            ereg = re.compile(
                '[1-9][0-9]{5}(19[0-9]{2}|20[0-9]{2})((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|1[0-9]|2[0-8]))[0-9]{3}[0-9Xx]$')  # //平年出生日期的合法性正则表达式
        # //测试出生日期的合法性
        if (re.match(ereg, id_code)):
            # //计算校验位
            S = (int(id_code_list[0]) + int(id_code_list[10])) * 7 + (
                    int(id_code_list[1]) + int(id_code_list[11])) * 9 + (
                        int(
                            id_code_list[2]) + int(id_code_list[12])) * 10 + (
                        int(id_code_list[3]) + int(id_code_list[13])) * 5 + (int(
                id_code_list[4]) + int(id_code_list[14])) * 8 + (int(id_code_list[5]) + int(id_code_list[15])) * 4 + (
                        int(
                            id_code_list[6]) + int(id_code_list[16])) * 2 + int(id_code_list[7]) * 1 + int(
                id_code_list[8]) * 6 + int(
                id_code_list[9]) * 3
            Y = S % 11
            M = "F"
            JYM = "10X98765432"
            M = JYM[Y]  # 判断校验位
            if (M == id_code_list[17]):  # 检测ID的校验位
                return Errors[0]
            else:
                return Errors[3]
        else:
            return Errors[2]
    else:
        return Errors[1]


# 多种格式的电话号码校验
'''
def isTelPhoneNumber(telphone):
    if len(telphone) == 11:
        if re.match(r"^1(?:749)\d{7}$", telphone):
            return 'MSC'
        elif re.match(r"^174(?:0[6-9]|1[0-2])\d{6}$", telphone):
            return 'MCC'
        elif re.match(r"^1(?:349)\d{7}$", telphone):
            return 'CM_SMC'
        elif re.match(r"^1(?:740[0-5])\d{6}$", telphone):
            return 'CT_SMC'
        elif re.match(r"^1(?:47)\d{8}$", telphone):
            return 'CM_IDC'
        elif re.match(r"^1(?:45)\d{8}$", telphone):
            return 'CU_IDC'
        elif re.match(r"^1(?:49)\d{8}$", telphone):
            return 'CT_IDC'
        elif re.match(r"^1(?:70[356]|65\d)\d{7}$", telphone):
            return 'CM_VNO'
        elif re.match(r"^1(?:70[4,7-9]|71\d|67\d)\d{7}$", telphone):
            return 'CU_VNO'
        elif re.match(r"^1(?:70[0-2]|62\d)\d{7}$", telphone):
            return 'CT_VNO'
        elif re.match(r"^1(?:34[0-8]|3[5-9]\d|5[0-2,7-9]\d|7[28]\d|8[2-4,7-8]\d|9[5,7,8]\d)\d{7}$", telphone):
            return 'CM_BO'
        elif re.match(r"^1(?:3[0-2]|[578][56]|66|96)\d{8}$", telphone):
            return 'CU_BO'
        elif re.match(r"^1(?:33|53|7[37]|8[019]|9[0139])\d{8}$", telphone):
            return 'CT_BO'
        elif re.match(r"^1(?:92)\d{8}$", telphone):
            return 'CBN_BO'
        else:
            return 'Error'
    elif len(telphone) == 13:
        if re.match(r"^14(?:40|8\d)\d{9}$", telphone):
            return 'CM_IoT'
        elif re.match(r"^14(?:00|6\d)\d{9}$", telphone):
            return 'CU_IoT'
        elif re.match(r"^14(?:10)\d{9}$", telphone):
            return 'CT_IoT'
        else:
            return 'Error'
    #add by Wzk
    #检查10位电话号码，只检查长度
    elif len(telphone) == 10:
        return "True"
    else:
        return 'Error'

    # if result == 'Error':
    #     print("你输入的号码不正确，请重新输入！")
    # elif result == 'MSC':
    #     print('你的号码是海事卫星通信的。')
    # elif result == 'MCC':
    #     print('你的号码是工信部应急通信的。')
    # elif result == 'CM_SMC':
    #     print('你的号码是中国移动卫星通信的。')
    # elif result == 'CT_SMC':
    #     print('你的号码是中国电信卫星通信的。')
    # elif result == 'CM_IDC':
    #     print('你的号码是中国移动上网数据卡的。')
    # elif result == 'CU_IDC':
    #     print('你的号码是中国联通上网数据卡的。')
    # elif result == 'CT_IDC':
    #     print('你的号码是中国电信上网数据卡的。')
    # elif result == 'CM_VNO':
    #     print('你的号码是中国移动虚拟运营商的。')
    # elif result == 'CU_VNO':
    #     print('你的号码是中国联通虚拟运营商的。')
    # elif result == 'CT_VNO':
    #     print('你的号码是中国电信虚拟运营商的。')
    # elif result == 'CM_BO':
    #     print('你的号码是中国移动的。')
    # elif result == 'CU_BO':
    #     print('你的号码是中国联通的。')
    # elif result == 'CT_BO':
    #     print('你的号码是中国电信的。')
    # elif result == 'CBN_BO':
    #     print('你的号码是中国广电的。')
    # elif result == 'CM_IoT':
    #     print('你的号码是中国移动物联网数据卡的。')
    # elif result == 'CU_IoT':
    #     print('你的号码是中国联通物联网数据卡的。')
    # elif result == 'CT_IoT':
    #     print('你的号码是中国电信物联网数据卡的。')
'''
def isTelPhoneNumber(a):
    a = a.replace(" ", "").replace("+", "").replace("-", "").replace("（","").replace("）","")
    b = re.match("^(?:\+?86)?1(?:3\d{3}|5[^4\D]\d{2}|8\d{3}|7(?:[235-8]\d{2}|4(?:0\d|1[0-2]|9\d))|9[0-35-9]\d{2}|66\d{2})\d{6}$",a)
    #b=re.match("^(?:\+?86)?1(?:3\d{3}|5[^4\D]\d{2}|8\d{3}|7(?:[0-35-9]\d{2}|4(?:0\d|1[0-2]|9\d))|9[0-35-9]\d{2}|6[2567]\d{2}|4(?:(?:10|4[01])\d{3}|[68]\d{4}|[579]\d{2}))\d{6}$",a)

    c = re.match("^0?(10|(2|3[1,5,7]|4[1,5,7]|5[1,3,5,7]|7[1,3,5,7,9]|8[1,3,7,9])[0-9]|91[0-7,9]|(43|59|85)[1-9]|39[1-8]|54[3,6]|(701|580|349|335)|54[3,6]|69[1-2]|44[0,8]|48[2,3]|46[4,7,8,9]|52[0,3,7]|42[1,7,9]|56[1-6]|63[1-5]|66[0-3,8]|72[2,4,8]|74[3-6]|76[0,2,3,5,6,8,9]|82[5-7]|88[1,3,6-8]|90[1-3,6,8,9])\d{7,8}$",a)
    #print(b)
    #print(c)
    if b != None or c != None:
        return True
    else:
        return "Error"


# 年月日,时间 是否规范，如 2018-02-30
def isRightDate(y, m, d):
    date = f'{y}-{m}-{d}'
    try:
        datetime.datetime.strptime(date, "%Y-%m-%d")
        return True
    except:
        return False


# 判断qq
def checkQQ(qq):
    pattern = r"[1-9]\d{4,6}"
    res = re.findall(pattern, qq, re.I)
    if len(res) == 1:
        return True
    else:
        return False


# 判断邮箱
'''
def checkEmail(eamil):
    pattern = r"\w{0,19}@[0-9a-zA-Z]{1,13}\.[com,cn,net]{1,3}"
    res = re.findall(pattern, eamil, re.I)
    if len(res) == 1:
        return True
    else:
        return False
'''
def checkEmail(eamil):
    pattern = '^[*#\u4e00-\u9fa5 a-zA-Z0-9_.-]+@[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)*\.[a-zA-Z0-9]{2,6}$'
    res = re.findall(pattern, eamil, flags=0)
    #print(res)
    if len(res) == 1:
        return True
    else:
        return False


def addRemarkInDoc(word, document, f, content):
    word.Selection.End = 0
    word.Selection.Start = 0
    word.Selection.Find.Execute(f)
    document.Comments.Add(Range=word.Selection.Range, Text=content)


def str_insert(s, index, sub):
    if sub in s:
        return
    s_list = list(s)
    s_list.insert(index, sub)
    s = "".join(s_list)
    return s

#add by suchao 邮箱检验
def isEmail(emailnumber):
    pattern = re.compile(r"[a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(?:\.[a-zA-Z0-9_-]+)")
    result =  pattern.findall(emailnumber)
    if result == []:
        return 'Error'
    else:
        return result

# add by suchao 判断是否包含小数点，还有判断是否只包含小数点和数字的这边没写出来,用于digital_to_Upper()
def is_contain_dot(check_str):
    check_str = str(check_str)
    for ch in check_str:
        if ch == '.':
            return True
    return False

#add by suchao 金额大小写检验
def digital_to_Upper(moneystr):
    nums = {0: '零', 1: '壹', 2: '贰', 3: '叁', 4: '肆', 5: '伍', 6: '陆', 7: '柒', 8: '捌', 9: '玖'}
    decimal_label = ['角', '分']
    small_int_label = {0: '', 1: '拾', 2: '佰', 3: '仟', 4: '万', 5: '拾', 6: '佰', 7: '仟', 8: '亿'}
    decimal_part = ''
    integer_part_list = []
    integer_part = ''
    # 包含小数点，则分成整数部分和小数部分
    if is_contain_dot(moneystr) is True:
        integer, decimal = str(moneystr).split('.', 1)
        if len(decimal) > 2:
            # print('小数部分超出')
            return 'Error'
        elif len(integer) > 9:
            return 'Error'
        # 处理小数部分，只处理到百分位
        for i, j in enumerate(decimal):
            """
            i: 记录循环次数
            """
            if j == '0' and decimal[-1] != '0':
                decimal_part += nums[int(j)]
            elif j == '0' and decimal[-1] == '0':
                pass
            else:
                decimal_part += (nums[int(j)] + decimal_label[i])
    # 不包含小数点，则为整数部分
    else:
        integer = str(moneystr)
        if len(integer) > 9:
            return 'Error'

    """
    处理整数部分,到亿；这边的处理办法是从低位往高位读，遇到0，判断前一位是否为0，前一位为0的情况则不读，如果前一位不为0，则读0。
    """
    # if integer != '' and int(integer) != 0:
    #     integer_part_list.insert(0, '元')
    for n, m in enumerate(integer[::-1]):

        if n == 0 and m == '0':
            pass
        else:
            # 当前为0 同时在万位，并且输入不为0时插入“万”
            if m == '0' and n == 4 and int(integer) != 0:
                integer_part_list.insert(0, '万')

            elif m == '0' and integer[::-1][n - 1] != '0':
                integer_part_list.insert(0, (nums[int(m)]))

            elif m == '0' and integer[::-1][n - 1] == '0':
                pass
            else:
                integer_part_list.insert(0, (nums[int(m)] + small_int_label[n]))
    integer_part = ''.join(integer_part_list)
    return (integer_part + decimal_part)
def check_str(re_exp,str):
    res=re.search(re_exp,str)
    if res:
        return True
    else:
        return False

#print(check_str('[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}','91510000201893845'))


#ADD BY QY
# 判断是否全是空格或者回车，全是返回False
def checkEntersAndSpace(s):
    for i in s:
        if i != '\n' and i != ' ':
            return True
    return False

#ADD by tyh
def is_youbian(number):
    pattern = r'[1-9][0-9]{5}'
    result = re.findall(pattern, number)
    if result == []:
        return False
    else:
        return result

def is_bankcard(number):
    number=number.replace(" ","")
    if len(number)==16 or len(number)==18 or len(number)==19 or len(number)==20:
        return True
    else:
        return False


def is_chuanzhen(number):
    pattern = r'(0\d{2,3})-)(\d{7,8})(-(\d{3,})'
    result = re.findall(pattern, number)
    if result == []:
        return False
    else:
        return result

def list_all_null(list):
    if list[0] == '':
        for t in list:
            if t!='':
                return False
        return True
    else:
        return False

def list_all_full(list) :
    if list[0] == '':
        return False
    else:
        for t in list:
            if t=='':
                return False
        return True


def table_ok(table):
    for row in table:
        if list_all_null(row)==False and list_all_full(row)==False:
             return False
    return True



def check_money(s):
    pattern = '（大写）(.*)（¥(.*)元）.*'
    l = re.findall(pattern, s)
    if l == [("", "")]:
        return False
    else:
        big = l[0][0]
        small = l[0][1]
        if rmbTrans.trans(big) == None:
            return False
        else:
            if float(rmbTrans.trans(big)) != float(small):
                return False
            else:
                return True

def check_money_split(s):
    pattern = '（大写）(.*)（¥(.*)元）.*'
    l = re.findall(pattern, s)
    if l == [("", "")]:
        return False
    else:
        big = l[0][0]
        small = l[0][1]
        return [big,small]


def check_money1(s):
    pattern = '（大写）(.*)元（¥(.*)）.*'
    l = re.findall(pattern, s)
    if l == [("", "")]:
        return False
    else:
        big = l[0][0].replace("圆", "").replace("元", "")
        small = l[0][1]
        if rmbTrans.trans(big) == None:
            return False
        else:
            if float(rmbTrans.trans(big)) != float(small):
                return False
            else:
                return True
