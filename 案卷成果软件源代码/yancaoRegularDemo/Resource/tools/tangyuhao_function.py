import os
import re
import datetime
import time
import winreg

from yancaoRegularDemo.Resource.ReadFile import DocxData

"""
    get_strtime(self, text)：得到 年-月-日 形式的时间
    is_id_number(id_number)：判断身份证号的正确性
    time_differ(time0,time1)：输出time0在time1后多少天
    sign_date(text)：返回特定格式下的 签名，日期
"""


def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                         r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', )
    return winreg.QueryValueEx(key, "Desktop")[0]


def subChar(str):
    match = re.compile(u'[\u4e00-\u9fa5]')
    return match.sub('', str)


def get_strtime_5_with_(text):
    text = text.replace("年", "-").replace("月", "-").replace("日", "-").replace("时", "-").replace("分", " ").replace("/",
                                                                                                                  "-").strip().replace(
        " ", "")
    text = re.sub("\s+", " ", text)
    t = ""
    regex_list = [
        "(\d{4}-\d{1,2}-\d{1,2}-\d{1,2}-\d{1,2})",

    ]
    for regex in regex_list:
        t = re.search(regex, text)
        if t:
            t = t.group(1)
            return t
    else:
        return False


def get_strtime_5(text):
    if "年" not in text or "月" not in text or "日" not in text or "时" not in text or "分" not in text:
        return False
    text = text.replace("年", "-").replace("月", "-").replace("日", "-").replace("时", "-").replace("分", " ").replace("/",
                                                                                                                  "-").strip().replace(
        " ", "")
    text = re.sub("\s+", " ", text)
    t = ""
    regex_list = [
        "(\d{4}-\d{1,2}-\d{1,2}-\d{1,2}-\d{1,2})",

    ]
    for regex in regex_list:
        t = re.search(regex, text)
        if t:
            t = t.group(1)
            return t
    else:
        return False


def get_strtime_with_(text):
    text = text.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip().replace(" ", "")
    text = re.sub("\s+", " ", text)
    t = ""
    regex_list = [
        # # 2013年8月15日 22:46:21
        # "(\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}:\d{1,2})",
        # # "2013年8月15日 22:46"
        # "(\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2})",
        # "2014年5月11日"
        "(\d{4}-\d{1,2}-\d{1,2})",

    ]
    for regex in regex_list:
        t = re.search(regex, text)
        if t:
            t = t.group(1)
            return t
    else:
        return False


def get_strtime(text):
    text = text.replace("号", "日")
    if "年" not in text or "月" not in text or "日" not in text:
        return False
    text = text.replace("年", "-").replace("月", "-").replace("日", " ").replace("/", "-").strip().replace(" ", "")
    text = re.sub("\s+", " ", text)
    t = ""
    regex_list = [
        # # 2013年8月15日 22:46:21
        # "(\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}:\d{1,2})",
        # # "2013年8月15日 22:46"
        # "(\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2})",
        # "2014年5月11日"
        "(\d{4}-\d{1,2}-\d{1,2})",

    ]
    for regex in regex_list:
        t = re.search(regex, text)
        if t:
            t = t.group(1)
            return t
    else:
        return False


def is_id_number(id_number):
    if len(id_number) != 18 and len(id_number) != 15:
        print('身份证号码长度错误')
        return False
    regularExpression = "(^[1-9]\\d{5}(18|19|20)\\d{2}((0[1-9])|(10|11|12))(([0-2][1-9])|10|20|30|31)\\d{3}[0-9Xx]$)|" \
                        "(^[1-9]\\d{5}\\d{2}((0[1-9])|(10|11|12))(([0-2][1-9])|10|20|30|31)\\d{3}$)"
    if re.match(regularExpression, id_number):
        if len(id_number) == 18:
            n = id_number.upper()
            # 前十七位加权因子
            var = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
            # 这是除以11后，可能产生的11位余数对应的验证码
            var_id = ['1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2']

            sum = 0
            for i in range(0, 17):
                sum += int(n[i]) * var[i]
            sum %= 11
            if (var_id[sum]) != str(n[17]):
                print("身份证号规则核验失败，校验码应为", var_id[sum], "，当前校验码是：", n[17])
                return False
        return True
    else:
        return False


def time_differ(time0, time1):
    date1 = time.strptime(time1, "%Y-%m-%d")
    date0 = time.strptime(time0, "%Y-%m-%d")
    date1 = datetime.datetime(date1[0], date1[1], date1[2])
    date0 = datetime.datetime(date0[0], date0[1], date0[2])
    return (date0 - date1).days

#大于0表示time0在time1前
def time_differ_5(time0, time1):
    date1 = time.strptime(time1, "%Y-%m-%d-%H-%M")
    date0 = time.strptime(time0, "%Y-%m-%d-%H-%M")
    d=(date1[0]-date0[0])*366*24*60+(date1[1]-date0[1])*31*24*60+(date1[2]-date0[2])*24*60+(date1[3]-date0[3])*60+date1[4]-date0[4]
    return d


def B2Q(uchar):
    """单个字符 半角转全角"""
    inside_code = ord(uchar)
    if inside_code < 0x0020 or inside_code > 0x7e:  # 不是半角字符就返回原来的字符
        return uchar
    if inside_code == 0x0020:  # 除了空格其他的全角半角的公式为: 半角 = 全角 - 0xfee0
        inside_code = 0x3000
    else:
        inside_code += 0xfee0
    return chr(inside_code)


def strB2Q(ustring):
    """半角转全角"""
    rstring = ""
    for uchar in ustring:
        inside_code = ord(uchar)
        if inside_code == 32:  # 半角空格直接转化
            inside_code = 12288
        elif inside_code >= 32 and inside_code <= 126:  # 半角字符（除空格）根据关系转化
            inside_code += 65248

        rstring += chr(inside_code)
    return rstring


def sign_date(text):
    text = text.replace(":", "：")
    pattern = re.compile(r'.*签名：(.*)日期.*')
    sign = re.findall(pattern, text)

    pattern = re.compile(r'.*日期：(.*)')
    date = re.findall(pattern, text)

    return sign, date


# f是打批注的地方
def addRemarkInDoc(word, document, f, content):
    word.Selection.End = 0
    word.Selection.Start = 0
    word.Selection.Find.Execute(f)
    document.Comments.Add(Range=word.Selection.Range, Text=content)
    document.Save()


def allSame_noSpace(list1):
    l = len(list1)
    if l == 0:
        return True
    x = list1[0].replace(" ", "")
    for i in range(0, l):
        if list1[i].replace(" ", "") != x:
            return False
    return True


def startTime(source_prifix):
    if os.path.exists(source_prifix + "立案报告表_.docx") == 1:
        data = DocxData(source_prifix + "立案报告表_.docx")
        form = data.tabels_content
        return get_strtime_5(form["案发时间"].replace(" ", ""))
    else:
        return False


def endTime(source_prifix):
    if os.path.exists(source_prifix + "结案报告表_.docx") == 1:
        data = DocxData(source_prifix + "结案报告表_.docx")
        form = data.tabels_content
        text = form["负责人意见"]
        pattern = r".*日期：(.*)"
        text = re.findall(pattern, text)
        if text == [] or text[0].strip() == "" or text[0].replace(" ", "") == '/':
            return False
        else:
            return get_strtime(text[0].replace(" ", ""))

    else:
        return False


def startPlace(source_prifix):
    if os.path.exists(source_prifix + "立案报告表_.docx") == 1:
        data = DocxData(source_prifix + "立案报告表_.docx")
        form = data.tabels_content
        if form["案发地点"].replace(" ", "") == "":
            return False
        else:
            return form["案发地点"].replace(" ", "")
    else:
        return False


def jiancha_time(source_prifix):
    if os.path.exists(source_prifix + "检查（勘验）笔录_.docx") == 1:
        data = DocxData(source_prifix + "检查（勘验）笔录_.docx")
        text = data.text
        pattern = ".*检查（勘验）时间：(.*)至(.*).*"
        x = re.findall(pattern, text)
        # if x[0][0].strip() == "" or x[0][1].strip == "":
        #     return False
        return x
    else:
        return False


def jiancha_place(source_prifix):
    if os.path.exists(source_prifix + "检查（勘验）笔录_.docx") == 1:
        data = DocxData(source_prifix + "检查（勘验）笔录_.docx")
        text = data.text
        pattern = ".*检查（勘验）地点：(.*).*"
        return re.findall(pattern, text)
    else:
        return False


def file_exists_open(source_prifix, file_name, read_func):
    if '~$' not in file_name:
        for root, dirs, files in os.walk(source_prifix):
            for f in files:
                if '~$' not in f:
                    if file_name in f:
                        if '撤销' not in file_name and '撤销' in f:
                            continue
                        elif '撤销' in file_name and '撤销' not in f:
                            continue
                        elif os.path.exists(source_prifix + f) == 1:
                            return read_func(source_prifix + f)
                else:
                    continue
        return False
    else:
        return False


def file_exists(source_prifix, file_name):
    if '~$' not in file_name:
        for root, dirs, files in os.walk(source_prifix):
            for f in files:
                if '~$' not in f:
                    if file_name in f:
                        if os.path.exists(source_prifix + f) == 1:
                            return True
                else:
                    continue
        return False
    else:
        return False


def list_str_match(list, str):
    for i, v in enumerate(list):
        if v not in str:
            return False
    return True


def read_txt(file):
    f = open(file, encoding='utf-8')
    list = []
    for line in f.readlines():
        line = line.strip('\n')  # 去掉列表中每一个元素的换行符
        list.append(line)
    return list


def changeDate(str0):
    CN_NUM = {
        '〇': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '零': 0,
        '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5, '陆': 6, '柒': 7, '捌': 8, '玖': 9, '貮': 2, '两': 2,
    }

    CN_UNIT = {
        '十': 10,
        '拾': 10,
        '百': 100,
        '佰': 100,
        '千': 1000,
        '仟': 1000,
        '万': 10000,
        '萬': 10000,
        '亿': 100000000,
        '億': 100000000,
        '兆': 1000000000000,
    }

    def chinese_to_arabic(cn: str) -> int:
        if cn.isdigit():
            return cn
        unit = 0  # current
        ldig = []  # digest
        for cndig in reversed(cn):
            if cndig in CN_UNIT:
                unit = CN_UNIT.get(cndig)
                if unit == 10000 or unit == 100000000:
                    ldig.append(unit)
                    unit = 1
            else:
                dig = CN_NUM.get(cndig)
                if unit:
                    dig *= unit
                    unit = 0
                ldig.append(dig)
        if unit == 10:
            ldig.append(10)
        val, tmp = 0, 0
        for x in reversed(ldig):
            if x == 10000 or x == 100000000:
                val += tmp * x
                tmp = 0
            else:
                tmp += x
        val += tmp
        return val

    str0 = str0.replace(" ", "")
    list1 = []

    pattern = r"(.*)年.*"
    nian = re.findall(pattern, str0)
    list1.append(int(nian[0].replace("〇", "0").replace("0", "0").replace("○", "0").replace("一", "1").replace("二", "2").replace("三", "3").replace("四", "4").replace("五", "5").replace("六", "6").replace("七", "7").replace("八", "8").replace("九", "9")))

    pattern = r"年(.*)月.*"
    yue = re.findall(pattern, str0)
    list1.append(chinese_to_arabic(yue[0]))

    pattern = r"月(.*)日.*"
    ri = re.findall(pattern, str0)
    list1.append(chinese_to_arabic(ri[0]))

    return str(list1[0]) + "-" + str(list1[1]) + "-" + str(list1[2])


def ch2num(chstr):
    chinese_number_dict = {'壹': 1, '柒': 7, '万': 10000, '叁': 3, '玖': 9, '贰': 2, '伍': 5, '捌': 8, '陆': 6, '拾': 10,
                           '仟': 1000, '肆': 4, '佰': 100, '零': 0, }
    not_in_decimal = "拾佰仟万点"

    def ch2round(chstr):
        no_op = True
        if len(chstr) >= 2:
            for i in chstr:
                if i in not_in_decimal:
                    no_op = False
        else:
            no_op = False
        if no_op:
            return ch2decimal(chstr)

        result = 0
        now_base = 1
        big_base = 1
        big_big_base = 1
        base_set = set()
        chstr = chstr[::-1]
        for i in chstr:
            if i not in chinese_number_dict:
                return None
            if chinese_number_dict[i] >= 10:
                if chinese_number_dict[i] > now_base:
                    now_base = chinese_number_dict[i]
                elif now_base >= chinese_number_dict["万"] and now_base < chinese_number_dict["亿"] and \
                        chinese_number_dict[i] > big_base:
                    now_base = chinese_number_dict[i] * chinese_number_dict["万"]
                    big_base = chinese_number_dict[i]
                elif now_base >= chinese_number_dict["亿"] and chinese_number_dict[i] > big_big_base:
                    now_base = chinese_number_dict[i] * chinese_number_dict["亿"]
                    big_big_base = chinese_number_dict[i]
                else:
                    return None
            else:
                if now_base in base_set and chinese_number_dict[i] != 0:
                    return None
                result = result + now_base * chinese_number_dict[i]
                base_set.add(now_base)
        if now_base not in base_set:
            result = result + now_base * 1
        return result

    def ch2decimal(chstr):
        result = ""
        for i in chstr:
            if i in not_in_decimal:
                return None
            if i not in chinese_number_dict:
                return None
            result = result + str(chinese_number_dict[i])
        return int(result)

    if '点' not in chstr:
        return ch2round(chstr)
    splits = chstr.split("点")
    if len(splits) != 2:
        return splits
    rount = ch2round(splits[0])
    decimal = ch2decimal(splits[-1])
    if rount is not None and decimal is not None:
        return float(str(rount) + "." + str(decimal))
    else:
        return None


def formatCurrency(currencyDigits):
    '''本函数旨在将数字化的金额（不含千分符）转化为中文的大写金额'''
    maximum_number = 99999999999.99
    cn_zero = "零"
    cn_one = "壹"
    cn_two = "贰"
    cn_three = "叁"
    cn_four = "肆"
    cn_five = "伍"
    cn_six = "陆"
    cn_seven = "柒"
    cn_eight = "捌"
    cn_nine = "玖"
    cn_ten = "拾"
    cn_hundred = "佰"
    cn_thousand = "仟"
    cn_ten_thousand = "万"
    cn_hundred_million = "亿"
    cn_symbol = "人民币"
    cn_dollar = "圆"
    cn_ten_cent = "角"
    cn_cent = "分"
    cn_integer = "整"
    integral = None
    decimal = None
    outputCharacters = None
    parts = None
    digits, radices, bigRadices, decimals = None, None, None, None
    zeroCount = None
    i, p, d = None, None, None
    quotient, modulus = None, None
    currencyDigits = str(currencyDigits)
    if currencyDigits == "":
        return ""
    if float(currencyDigits) > maximum_number:
        print("转换金额过大!")
        return ""
    parts = currencyDigits.split(".")
    if len(parts) > 1:
        integral = parts[0]
        decimal = parts[1]
        decimal = decimal[0:2]
        if decimal == "0" or decimal == "00":
            decimal = ""
    else:
        integral = parts[0]
        decimal = ""
    digits = [cn_zero, cn_one, cn_two, cn_three, cn_four, cn_five, cn_six, cn_seven, cn_eight, cn_nine]
    radices = ["", cn_ten, cn_hundred, cn_thousand]
    bigRadices = ["", cn_ten_thousand, cn_hundred_million]
    decimals = [cn_ten_cent, cn_cent]
    outputCharacters = ""
    if int(integral) > 0:
        zeroCount = 0
        for i in range(len(integral)):
            p = len(integral) - i - 1
            d = integral[i]
            quotient = int(p / 4)
            modulus = p % 4
            if d == "0":
                zeroCount += 1
            else:
                if zeroCount > 0:
                    outputCharacters += digits[0]
                zeroCount = 0
                outputCharacters = outputCharacters + digits[int(d)] + radices[modulus]
            if modulus == 0 and zeroCount < 4:
                outputCharacters = outputCharacters + bigRadices[quotient]
        outputCharacters += cn_dollar
    if decimal != "":
        jiao = decimal[0]
        if jiao == "":
            jiao = "0"
        try:
            fen = decimal[1]
        except:
            fen = "0"
        if fen == "":
            fen = "0"
        if jiao == "0" and fen == "0":
            pass
        if jiao == "0" and fen != "0":
            outputCharacters = outputCharacters + cn_zero + digits[int(fen)] + decimals[1]
        if jiao != "0" and fen == "0":
            outputCharacters = outputCharacters + digits[int(jiao)] + decimals[0]
        if jiao != "0" and fen != "0":
            outputCharacters = outputCharacters + digits[int(jiao)] + decimals[0]
            outputCharacters = outputCharacters + digits[int(fen)] + decimals[1]
    if outputCharacters == "":
        outputCharacters = cn_zero + cn_dollar
    if decimal == "":
        outputCharacters = outputCharacters + cn_integer
    outputCharacters = outputCharacters
    return outputCharacters


# for currency in [2640.80]:
#      capital_currency=formatCurrency(currency)
#      print(str(currency)+":\t"+capital_currency)


if __name__ == "__main__":
    string = "ssss，ssss,sss"
    print(subChar(string))
