class Money(object):#输入为阿拉伯，或者大小写的数字,字符格式
    def __init__(self):
        self.chinese=['一','二','三','四','五','六','七','八','九','十','百','千','万']
        self.big=['壹','贰','叁','肆','伍','陆','柒','捌','玖','拾','佰','陌','仟','阡','萬']
        self.CN_NUM = {
            '〇': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '零': 0,
            '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5, '陆': 6, '柒': 7, '捌': 8, '玖': 9, '貮': 2, '两': 2,
            '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9, '0': 0
        }
        self.CN_UNIT = {
            '十': 10,
            '拾': 10,
            '百': 100,
            '佰': 100,
            '千': 1000,
            '仟': 1000,
            '阡':1000,
            '万': 10000,
            '萬': 10000,
            '亿': 100000000,
            '億': 100000000,
            '兆': 1000000000000,
        }
    def chinese_to_arabic(self,cn: str) -> int:
        unit = 0  # current
        ldig = []  # digest
        for cndig in reversed(cn):
            if cndig in self.CN_UNIT:
                unit = self.CN_UNIT.get(cndig)
                if unit == 10000 or unit == 100000000:
                    ldig.append(unit)
                    unit = 1
            else:
                dig = self.CN_NUM.get(cndig)
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
    def form(self,money):
        chinese=0
        big=0
        arab=0
        for i in money:
            if i not in self.chinese:
                chinese=1   #标志位设为1表示不是chinese类的字符
            if i not in self.big:
                big=1
            if  not i.isdigit():
                arab=1
        if chinese==0 or big==0:
            rst=self.chinese_to_arabic(money)
        else:rst=money
        if chinese==0:
            return int(rst),1  #第二个数为1代表输入的金额为简体的形式
        elif big==0:
            return int(rst),2  #第二个数为2代表输入的金额为繁体的形式
        elif arab==0:
            return int(rst),3   #第二个数为3代表输入的金额为阿拉伯的形式
        else:
            return 0,0          #返回双0代表判断出错
    def fit(self,money1,money2):
        money1,_=self.form(money1)
        money2,_=self.form(money2)
        if money2==money1:
            return 1   #返回1表示一致
        else:return 0  #返回0表示不一致
    def penalty(self,money1,money2):  #第一个为总金额 第二个为罚款金额
        money1, _ = self.form(money1)
        money2, _ = self.form(money2)
        money1 = money1*0.1
        if money2<=money1:
            return 1         #返回1表示合法
        else:return 0       #返回2表示不合法
    def num_juge(self,money,num):
        money,_=self.form(money)
        if money>num:
            return 1   #返回1表示金额大于数值
        elif money==num:
            return 2   #返回2表示金额等于数值
        else:return 3 #返回3表示金额小于数值




