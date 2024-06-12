"""
    Project:中国烟草案卷执法组身份证文本处理类
    Author:谢俊宇
    Date:2022-02-28 17:26
"""
import datetime
import re
from cid import IdParser

class IdCard_Information_Processor(object):

    def __init__(self, id):
        self.id = id
        self.birth_year = int(self.id[6:10])
        self.birth_month = int(self.id[10:12])
        self.birth_day = int(self.id[12:14])

    def is_valid_cid(self):
        """校验身份证号码格式是否正确"""
        ip = IdParser
        return ip.is_valid_cid(self.id)

    def get_region(self):
        """提取发证地"""
        ip = IdParser
        return ip.extract_region(self.id)

    def get_birthday(self):
        """通过身份证号获取出生日期"""
        birthday = "{0}-{1}-{2}".format(self.birth_year, self.birth_month, self.birth_day)
        return birthday

    def get_sex(self):
        """男生：1 女生：2"""
        num = int(self.id[16:17])
        if num % 2 == 0:
            return 2
        else:
            return 1

    def get_age(self):
        """通过身份证号获取年龄"""
        now = (datetime.datetime.now() + datetime.timedelta(days=1))
        year = now.year
        month = now.month
        day = now.day

        if year == self.birth_year:
            return 0
        else:
            if self.birth_month > month or (self.birth_month == month and self.birth_day > day):
                return year - self.birth_year - 1
            else:
                return year - self.birth_year


# idcard_string = '身份证：511622198308017318'
# idcard_number = re.search(r'([1-9]\d{5}[12]\d{3}(0[1-9]|1[012])(0[1-9]|[12][0-9]|3[01])\d{3}[0-9xX])',idcard_string,re.S)
# if idcard_number is not None:
#     print(idcard_number.group())
# print(idcard_number)