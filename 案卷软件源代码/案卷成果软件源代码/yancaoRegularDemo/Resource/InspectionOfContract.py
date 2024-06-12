# -*- coding:utf-8 -*-
# @ModuleName: InspectionOfContract
# @Function: 
# @Author: huhonghui、chenfumin、leitianyi
# @email: 1241328737@qq.com
# @Time: 2021/5/18 19:10
from yancaoRegularDemo.Resource.ReadFile import DocxData
import re
from yancaoRegularDemo.Resource.tools.utils import *


class InspectionOfContract:
    def __init__(self):
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.price_basis()",
            "self.uniformity_of_locations()",
            "self.uniformity_of_causes()",
            "self.uniformity_of_loc_date()",
            "self.uniformity_of_sign_name_date()",
            "self.uniformity_of_case_happen_date()",
            "self.uniformity_of_fine_money()",
        ]

    def _display(self,text,color = None):
        if color == 'red':
            text = "\033[0;31m" + "warning: " + text + "\033[0m"
            print(text)
        elif color == "green":
            text = "\033[0;36m" + text + "\033[0m"
            print(text)
        else:
            print(text)

    def price_basis(self):
        '''
        作用：检查是否包含涉案物品核价依据或价格来源
        :return: True for "包含"
                 False for "不包含"
        '''
        result = re.findall("依据(.*?)计算",self.contract_text)
        if len(result) > 0 :
            self._display("√ 包含涉案物品核价依据或价格来源:"+result[0], "green")
            return result[0]
        self._display("不包含涉案物品核价依据或价格来源", "red")
        return None

    def uniformity_of_locations(self):
        '''
        作用：检查文案前后案发地点的一致性
        :return: True for "一致"
                 False for "不一致"
        '''
        location_in_tables = self.contract_tables_content["案发地点"]
        if not isinstance(location_in_tables, list):
            location_in_tables = [location_in_tables]
        location_in_text = re.findall("在(.*?)查获",self.contract_text)    #目前文书数量只有1份，可能筛选条件不够宽泛，鲁棒性不强

        # print(location_in_tables,location_in_text)

        for i in location_in_tables:
            for j in location_in_text:
                if i not in j:
                    self._display('文案前后案发地点不一致\"' + i +  '\" 与 \"' + j + '\" 矛盾！',"red")
                    return None
        self._display("√ 文案前后案发地点一致:" + location_in_tables[0], "green")
        return location_in_tables[0]

    def uniformity_of_causes(self):
        '''
        作用：检查文案前后案由的一致性
        :return: True for "一致"
                 False for "不一致"
        备注：“案由”、“移送案由”的映射关系是什么？需要明确。例如此处的 无烟草专卖品准运证运输烟草专卖品==非法经营罪？？
        '''
        causes_in_tables = self.contract_tables_content["案由"]
        if not isinstance(causes_in_tables, list):
            causes_in_tables = [causes_in_tables]
        # print(causes_in_tables)

        causes_in_tables.sort(key = lambda i:len(i),reverse=False)
        for i in causes_in_tables[1:]:
            if causes_in_tables[0] not in i:
                self._display('文案前后案由不一致:\"' + causes_in_tables[0] +  '\" 与 \"' + i + '\" 矛盾！',"red")
                return None
        self._display("√ 案由一致：" + causes_in_tables[0], "green")
        return causes_in_tables[0]

    def uniformity_of_loc_date(self):
        '''
        作用：检查案发日期是否为空且格式是否正确,判断案发地点是否为空。
        '''
        case_date = self.contract_tables_content['案发时间']
        case_loc = self.contract_tables_content['案发地点']
        if len(case_date) < 1:
            self._display("× 案发时间没有填写！请重新核查！", "red")
        else:
            if not is_valid_date(case_date):
                self._display("× 案发时间不合法！时间格式为：XXXX年XX月XX日XX时XX分！请统一格式！文案日期为：" + case_date, "red")
            else:
                self._display("√ 案发时间输入合法！", "green")
        if len(case_loc) < 1:
            self._display("× 案发地点没有填写！请重新核查！", "red")
        else:
            self._display("√ 案发地点输入合法！", "green")

    def uniformity_of_sign_name_date(self):
        """
        作用：# 判断规则：没有承办人及承办部门负责人签字，没有签署日期的。
        :return: 有几处没有签字或者没有填写日期。
        """
        # 定义一个变量记录一共有多少处承办人相关的
        sum_mistakes_sign = 0
        sum_mistakes_date = 0
        sign_list = list()
        for key in ['承办人意见', '承办部门负责人意见', '承办部门意见']:
            if isinstance(self.contract_tables_content[key], list):
                sign_list.extend(self.contract_tables_content[key])
            elif isinstance(self.contract_tables_content[key], str):
                sign_list.append(self.contract_tables_content[key])
        for sign in sign_list:
            sign_index = sign.find('签名：')
            date_index = sign.find('日期')
            name = sign[sign_index+3: date_index]
            if len(name) != 0:
                pass
            else:
                sum_mistakes_sign += 1
            date = re.search('\d{4}年\d{1,2}月\d{1,2}日', sign)
            if date:
                pass
            else:
                sum_mistakes_date += 1
        if sum_mistakes_sign > 0:
            self._display('× 承办人相关部门有{}处未进行签字！请确认！'.format(sum_mistakes_sign), 'red')
        if sum_mistakes_date > 0:
            self._display('× 承办人相关部门有{}处未进行日期的填写！请确认！'.format(sum_mistakes_date), 'red')
        if sum_mistakes_date == 0 and sum_mistakes_sign == 0:
            self._display('√ 承办人相关部门都已进行签字和日期！', 'green')

    def uniformity_of_case_happen_date(self):
        """
        作用：没有记载或错误记载发案时间的（文案的前后发案时间要一致，不一致要提示）
        :return:
        """
        caseHappenDate = list()
        for key in ['案发时间', '案情摘要']:
            if isinstance(self.contract_tables_content[key], list):
                for content in self.contract_tables_content[key]:
                    date = re.search('\d{4}年\d{1,2}月\d{1,2}日', content).group(0)
                    caseHappenDate.append(date)
            elif isinstance(self.contract_tables_content[key], str):
                date = re.search('\d{4}年\d{1,2}月\d{1,2}日', self.contract_tables_content[key]).group(0)
                caseHappenDate.append(date)
            caseHappenDate = list(set(caseHappenDate))
        if len(caseHappenDate) == 1:
            self._display('√ 发案时间前后文一致！', 'green')
        else:
            self._display('× 案发时间和案情摘要中的案发时间不一致！请审查！', 'red')

    def uniformity_of_fine_money(self):
        """
        作用：判断是否有罚款金额
        :return:
        """
        content = '罚款，计金额：(.*?)元'
        number = re.compile(content)
        pattern = re.compile(number)
        all = pattern.findall(self.contract_text.replace('\n', ''))
        all = list(set(all))
        if len(all) == 0:
            self._display('× 不存在罚款金额，请核实！', 'red')
        elif len(all) == 1:
            self._display('√ 存在罚款金额，且无误！', 'green')
        else:
            self._display('× 前后文罚款金额不一致，请核实！', 'red')


    def check(self, contract_file_path):
        print("正在审查合同文件：%s....\n审查结果如下："%contract_file_path)
        data = DocxData(file_path=contract_file_path)
        self.contract_text = data.text
        # print('段落文本内容：', self.contract_text)
        self.contract_tables_content = data.tabels_content
        for func in self.all_to_check:
            eval(func)
        print("合同文件“%s”审查完毕\n"%contract_file_path)


if __name__ == '__main__':
    ioc = InspectionOfContract()
    contract_file_path = "data/"
    for file in os.listdir(contract_file_path):
        ioc.check(contract_file_path + file)
