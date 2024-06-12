# by wzk
# this file includes functions used in getFactorsFromContract
import copy
import datetime


# 合同金额大小写判断
# 金额小写转大写
def money_en_to_cn(value):
    """
    人民币大写
    来自：http://topic.csdn.net/u/20091129/20/b778a93d-9f8f-4829-9297-d05b08a23f80.html
    传入浮点类型的值返回 unicode 字符串
    """
    map = [u"零", u"壹", u"贰", u"叁", u"肆", u"伍", u"陆", u"柒", u"捌", u"玖"]
    # unit = [u"分", u"角", u"圆", u"拾", u"百", u"仟", u"万", u"拾", u"百", u"仟", u"亿",
    # u"拾", u"百", u"千", u"万", u"拾", u"百", u"千", u"兆"]
    unit = [u"分", u"角", u"圆", u"拾", u"佰", u"仟", u"万", u"拾", u"百", u"仟", u"亿",
            u"拾", u"佰", u"千", u"万", u"拾", u"百", u"千", u"兆"]

    nums = []  # 取出每一位数字，整数用字符方式转换避大数出现误差
    for i in range(len(unit) - 3, -3, -1):
        if value >= 10 ** i or i < 1:
            nums.append(int(round(value / (10 ** i), 2)) % 10)

    words = []
    zflag = 0  # 标记连续0次数，以删除万字，或适时插入零字
    start = len(nums) - 3
    for i in range(start, -3, -1):  # 使i对应实际位数，负数为角分
        if 0 != nums[start - i] or len(words) == 0:
            if zflag:
                words.append(map[0])
                zflag = 0
            words.append(map[nums[start - i]])
            words.append(unit[i + 2])
        elif 0 == i or (0 == i % 4 and zflag < 3):  # 控制‘万/元’
            words.append(unit[i + 2])
            zflag = 0
        else:
            zflag += 1

    # if words[-1] != unit[0]:  # 结尾非‘分’补整字
    # words.append(u"整")
    # print(words)
    return ''.join(words)


# 日期检查
def CheckDate(y, m, d):
    date = f'{y}-{m}-{d}'
    try:
        datetime.datetime.strptime(date, "%Y-%m-%d")
        return True
    except:
        return False


# 送货方式检查
def DeliveryType(type):
    if type == "自提" or type == "包送包卸货":
        return "交货"
    elif type == "包送包安装":
        return "安装完毕"
    else:
        return None


# 批注功能实现
def addRemarkInDoc(word, document, f, content, start=0, end=0):
    word.Selection.End = end
    # start=word.Selection.StartOf
    word.Selection.Start = start
    word.Selection.Find.Execute(f)
    document.Comments.Add(Range=word.Selection.Range, Text=content)


def addRemarkInDoc_Highlightened(word, document, f, content, ):
    try:
        word.Selection.End = 0
        # start=word.Selection.StartOf
        word.Selection.Start = 0
        word.Selection.Find.Execute(f)
        # 替换批注项的颜色
        word.Selection.Range.HighlightColorIndex = 7  # 替换背景颜色为白色 8,绿色11 ,yelow 7
        # word.Selection.Font.Color = 255  # 替换文字颜色为红色 255
        document.Comments.Add(Range=word.Selection.Range, Text=content)
    except:
        print(f, "出错")


# 传入待批注的语句列表
def Formal_Remark_list(in_list):
    1


def str_insert(s, index, sub):
    if sub in s:
        return
    s_list = list(s)
    s_list.insert(index, sub)
    s = "".join(s_list)
    return s
# money=money_en_to_cn(19000)
# # str="壹佰贰拾叁圆贰角"
# print(money)
