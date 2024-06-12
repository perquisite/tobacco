import re
from strsimpy.levenshtein import Levenshtein
levenshtein = Levenshtein()

class Simple_Content:
    def __init__(self):
        #预设的一些模式串
        self.pattern_strings={
            "license_plate_number": r'([京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领]{1}[A-Z]{1}(([A-HJ-NP-Z0-9]{5}[DF]{1})|([DF]{1}[A-HJ-NP-Z0-9]{5})))|([京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领]{1}[A-Z]{1}[A-HJ-NP-Z0-9]{4}[A-HJ-NP-Z0-9挂学警港澳]{1})'
            ,"phone_number": r"1(?:34[0-8]|3[5-9]\d|5[0-2,7-9]\d|7[28]\d|8[2-4,7-8]\d|9[5,7,8]\d|3[0-2]\d|[578][56]\d|66\d|96\d|33\d|53\d|7[37]\d|8[019]\d|9[0139]\d|92\d)\d{7}"
            ,
        }

    def is_null(self,content):#判断是否为空
        content="" if content is None else content
        content="".join(content.split())#除去可见与不可见的空白符
        num=0
        for i in content:
            if not i.isprintable():
                num+=1
        if num==len(content):#全是不可见字符，或者已经变成空串
            return True
        else:
            return False


    def match_re(self,pattern:str,string:str,flag=0):#返回string中匹配模式串pattern的所有结果
        return [(i.group(),i.span()) for i in re.finditer(pattern,string,flag)]


    def is_consistent(self,pre:str,now:str,strict:bool=True,fault_tolerance=0):#判断是内容否一致
        if strict:#严格匹配
            return pre==now
        else:#使用编辑距离
            return levenshtein.distance(pre,now)<=fault_tolerance#编辑距离小于指定值算一致



































