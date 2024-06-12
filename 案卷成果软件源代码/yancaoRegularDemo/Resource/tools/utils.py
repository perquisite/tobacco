import datetime
import time
import os
from datetime import date

# Define a fake pkg_resources that has the structure needed so as not
# not to raise an error in workalendar's __init__
class pkg_resources(object):
    def get_distribution(package):
        return Foo
class Foo:
    version = None
# inject it into the module cache so that it is found rather than the original
import sys
sys.modules["pkg_resources"] = pkg_resources
from workalendar.asia import China
# Clear up the mess we've made incase we need the real pkg_resources for something else.
del sys.modules["pkg_resources"]
# from workalendar.asia import China


def is_valid_date(strdate):
    # 判断是否是一个有效的日期字符串
    try:
        if "年" in strdate:
            time.strptime(strdate, "%Y年%m月%d日%H时%M分")
        else:
            time.strptime(strdate, "%Y年%m月%d日%H时%M分")
        return True
    except:
        return False


def get_root_dir():
    # 获取当前文件的目录
    cur_path = os.path.abspath(os.path.dirname(__file__))
    # 获取根目录
    root_path = cur_path[:cur_path.find("tobacco\\") + len("tobacco\\")]
    return str(root_path)


def get_nearest_working_day(date_input):
    # 获得输入的年-月-日格式的日期，往后顺延时，最接近的一个工作日，若输入日期本身为工作日则返回本身
    cal = China()
    # 如果是假期，将日期加一天，
    while not cal.is_working_day(date_input):
        date_input = date_input + datetime.timedelta(days=1)
    return str(date_input)


def chinese_to_date(chineseStr):
    # 将中文的年月日 转为 数字的年-月-日
    # example input：二〇二〇年九月二十四日 output：2020-9-24
    strch1 = '0一二三四五六七八九十'
    strch2 = '〇一二三四五六七八九十'
    y, m, d = '', '', ''
    if chineseStr.find('年') > 1:
        y = chineseStr[0:chineseStr.index('年')]
    if chineseStr.find('月') > 1:
        m = chineseStr[chineseStr.index('年') + 1:chineseStr.index('月')]
    if chineseStr.find('日') > 1:
        d = chineseStr[chineseStr.index('月') + 1:chineseStr.index('日')]
    # 年
    if len(y) == 4:
        if y.find('0') == 1:
            y = str(strch1.index(y[0:1])) + str(strch1.index(y[1:2])) + str(strch1.index(y[2:3])) + str(
                strch1.index(y[3:4]))
        else:
            y = str(strch2.index(y[0:1])) + str(strch2.index(y[1:2])) + str(strch2.index(y[2:3])) + str(
                strch2.index(y[3:4]))
    else:
        return None
    # 月
    if len(m) == 1:
        m = str(strch1.index(m))
    elif len(m) == 2:
        m = str(strch1.index(m[0:1]))[0:1] + str(strch1.index(m[1:2]))

    # 日
    if len(d) == 1:
        d = str(strch1.index(d))
    elif len(d) == 2:
        if len(str(strch1.index(d[0:1]))) == 1:
            d = str(strch1.index(d[0:1])) + str(strch1.index(d[1:2]))[1:2]
        else:
            d = str(strch1.index(d[0:1]))[0:1] + str(strch1.index(d[1:2]))
    elif len(d) == 3:
        d = str(strch1.index(d[0:1])) + str(strch1.index(d[2:3]))
    # 生成 日期
    if y != '' and m != '' and d != '':
        return y + '-' + m + '-' + d  # datetime.date(int(y), int(m), int(d))
    elif y != '' and m != '':
        return y + '-' + m  # datetime.date(int(y), int(m))
    elif y != '':
        return y
