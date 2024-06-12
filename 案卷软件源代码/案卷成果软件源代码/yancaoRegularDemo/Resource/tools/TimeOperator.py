"""
    Project:中国烟草案卷执法组时间操作类
    Author:陈付旻
    Date:2021-06-24 10:12
"""
import datetime
import re
import time
# pip install chinesecalendar
from chinese_calendar import is_holiday, is_workday


"""
注：获取到的日期一律统一为XXXX年XX月XX日
"""
class TimeOper(object):
    def __init__(self):
        super(TimeOper, self).__init__()

    # 判断某一个日期是否符合XXXX年XX月XX日的格式，例如2021年06月24日
    def is_valid_date(self, sourceTime):
        try:
            time.strptime(sourceTime, "%Y年%m月%d日")
            return True
        except:
            return False

    # 判断两个日期的先后顺序以及时间差
    def time_order(self, sourceTime, targetTime):
        patten = re.compile('\d+')
        result_source = patten.findall(sourceTime)
        result_targe = patten.findall(targetTime)
        source_list = list()
        target_list = list()
        for source in result_source:
            if source[:1] == '0':
                source = source[1:]
            source_list.append(int(source))
        for target in result_targe:
            if target[:1] == '0':
                target = target[1:]
            target_list.append(int(target))
        d1 = datetime.date(source_list[0], source_list[1], source_list[-1])
        d2 = datetime.date(target_list[0], target_list[1], target_list[-1])
        days_gap = (d1 - d2).days
        return days_gap  # source减去target的天数
        # if days_gap > 0:
        #     print('source在target之后，相差{}天'.format(days_gap))
        # elif days_gap == 0:
        #     print('source和target是同一天')
        # else:
        #     print('source在target之前，相差{}天'.format(-days_gap))

    # 判断两个日期是否一致
    def isSameDate(self, sourceDate, targetDate):
        if sourceDate == targetDate:
            print('两个日期一致！')
        else:
            print('两个日期不一致！')

    # 判断日期是否为节假日期
    def isHoliday(self, currentDate):
        patten = re.compile('\d+')
        result_source = patten.findall(currentDate)
        date = ''
        for source in result_source:
            date += source+'-'
        print(date[:-1])
        currentDate = datetime.datetime.strptime(date[:-1], '%Y-%m-%d')
        flag = is_holiday(currentDate)
        print(flag)

    # 获取当前日期，返回XXXX-XX-XX
    def getLocalDate(self):
        t = time.strftime("%Y-%m-%d", time.localtime())
        return t
