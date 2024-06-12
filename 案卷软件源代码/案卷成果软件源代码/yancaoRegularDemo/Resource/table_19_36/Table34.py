import math
import re
from time import sleep

from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.EntityRecognition import EntityRecognition
from yancaoRegularDemo.Resource.tools.IdCard_Information_Processor import IdCard_Information_Processor
from yancaoRegularDemo.Resource.tools.OCR_IDCard import OCR_IDCard
from zhishitupu.anjuan_zyclq import *

function_description_dict = {
    'check_cause_of_action': '案由应当与《案件调查终结报告》之后的全部法律文书中“案由”的记载一致。',
    'check_time': '立案日期应当与《立案报告表》中的“负责人意见”的时间一致。',
    'check_party_info': '当事人：“姓名、性别、民族、证件类型及号码、住址”信息应当与《证据复制（提取）单》中当事人身份证明中记载的信息一致。',
    'check_case_facts': '案件事实：应当包括时间、地点、查获卷烟的品种、数量、金额、违法事实等。'
                        '其中“时间”、“地点”应当与《立案报告表》中记载的一致，“品种”、“数量”应当与《物品清单》记载的一致。'
                        '“金额”应当与《涉案烟草专卖品核价表》中的合计金额一致。违法事实应针对不同的案由设置不同的审查规则。',
    'check_opinions_of_the_undertaker': '承办人意见：应当与《案件集体讨论记录》中的“结论性意见”基本一致。“签名”应当在2人以上。“日期”应当在《案件集体讨论记录》的时间之后。',
    'check_opinions_of_the_department': '承办部门意见：检索关键词“同意”、“不同意”，必须有其中一个。签名应当完整。日期应在“承办人意见”的日期之后。',
    'check_opinions_of_the_law_department': '法制部门意见：检索关键词“同意”、“不同意”，必须有其中一个。签名应当完整。日期应在“承办人意见”的日期之后。',
    'check_opinions_of_the_director': '负责人意见：必须检索出“同意”或“不同意”。日期应当在法制部门意见日期之后，签名完整。',
    'check_additional_function_one_ner': '1.日期应与检查（勘验）时间的开始时间一致；2.地点应与检查（勘验）地点一致；3.事件事实文本中是否含有执法人员要素',
    'check_additional_funtion_two_ner': '案件事实的无规则文本中是否含有烟草专卖零售许可证、卷烟数目、法律条款等要素',
    'about_discretionary_power': '自由裁量权加载出错',
}


def _not_empty(s):
    return s and s.strip()


# 案件处理审批表
class Table34(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prifix = source_prifix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.mw.Visible = 0
        self.mw.DisplayAlerts = 0
        # sleep(0.5)
        self.contract_text = None
        self.contract_tables_content = None
        self.entityrecognition = EntityRecognition()
        self.file_name_real = ""
        # self.target_file_name_list = [
        #     ['立案报告表_'],
        #     ['涉案物品返回清单_'],
        #     ['涉案烟草专卖品核价表_'],
        #     ['案件集体讨论记录_'],
        #     ['检查（勘验）笔录_']
        # ]

        self.all_to_check = [
            "self.check_cause_of_action()",
            "self.check_time()",
            "self.check_party_info()",
            "self.check_case_facts()",
            "self.check_opinions_of_the_undertaker()",
            "self.check_opinions_of_the_department()",
            "self.check_opinions_of_the_law_department()",
            "self.check_opinions_of_the_director()",
            "self.check_additional_function_one_ner()",
            "self.check_additional_funtion_two_ner()",
            "self.about_discretionary_power()"
        ]

    def about_discretionary_power(self):
        info_one, info_two = self.get_anyou()
        print(info_one)
        print(info_two)
        content = discretionary_power(info_one, info_two)
        if content:
            table_father.display(self, "自由裁量权：" + content, "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办人\r意见', content)

    # def check_proportion(self):
    #     # temp_list = []
    #     data = DocxData(self.source_prifix + "/案件处理审批表_.docx")
    #     # 文书中目前发现有两种表达罚款比例的形式
    #     temp_1 = re.search(r"(\d+)%以上(\d+)%以下", data.tabels_content["承办人意见"])
    #     temp_2 = re.search(r"(\d+)%-(\d+)%", data.tabels_content["承办人意见"])
    #     # 提取实际罚款比例
    #     temp_3 = re.search(r"(\d+).(\d+)元(\d+)%的罚款，计罚款(\d+).(\d+)元", data.tabels_content["承办人意见"])
    #     if temp_1 and temp_2 and temp_3:
    #         ratio = float(temp_3[3])
    #         # print('ratio', float(temp_3[3]))
    #         money = temp_3[1] + '.' + temp_3[2]
    #         # print('money', money)
    #         found = temp_3[4] + '.' + temp_3[5]
    #         # print('found', found)
    #         shouldFound = (float(ratio) / 100) * float(money)
    #         shouldFound = round(shouldFound, 2)
    #         found1 = round(float(found), 2)
    #         # print(shouldFound, found1, found1==shouldFound)
    #         if found1 != shouldFound:
    #             table_father.display(self, f"实际罚款金额有误，应罚款{shouldFound},实际罚款{found1}！", "red")
    #             content = f"实际罚款金额有误，应罚款{shouldFound},实际罚款{found1}！"
    #             loc1 = f'计罚款{found}元'
    #             tyh.addRemarkInDoc(self.mw, self.doc, loc1, content)
    #         if temp_1:
    #             temp_list = [temp_1.group(1), temp_1.group(2)]
    #             loc = f"{temp_1.group(1)}%以上{temp_1.group(2)}%以下"
    #         elif temp_2 is not None:
    #             temp_list = [temp_2.group(1), temp_2.group(2)]
    #             loc = f"{temp_2.group(1)}%-{temp_2.group(2)}%"
    #         else:
    #             temp_list = [0, 0]
    #         if temp_list == [0, 0]:
    #             pass
    #         elif ratio > float(temp_list[1]) or ratio < float(temp_list[0]):
    #             table_father.display(self, "× 实际罚款比例与法定罚款区间不一致！", "red")
    #             content = f"实际罚款比例{ratio}%与法定罚款区间{temp_list[0]}%-{temp_list[1]}%不一致！"
    #             tyh.addRemarkInDoc(self.mw, self.doc, loc, content)
    #     else:
    #         table_father.display(self, "× 未获取到罚款比例及金额的信息", "red")
    #         tyh.addRemarkInDoc(self.mw, self.doc, '承办人意见：', '未获取到罚款比例及金额的信息')

    def check_cause_of_action(self):
        """
        作用：案由：应当与《案件调查终结报告》之后的全部法律文书中“案由”的记载一致。
        """
        # a = self.contract_tables_content
        this_cause_of_action = self.contract_tables_content['案由']

        # 针对案由在表格中的法律文书
        file_with_table_list = ['案件调查终结报告', '延长调查终结审批表']

        for file_name in file_with_table_list:
            if not tyh.file_exists(self.source_prifix, file_name):
                table_father.display(self, "文件缺失：《" + file_name + "》.docx不存在", "red")
            else:
                file_info = tyh.file_exists_open(self.source_prifix, file_name, DocxData)
                if file_info is None:
                    table_father.display(self,
                                         "文件读取失败：" + file_name + "不存在，无法与《案件集体讨论记录》进行案由对比")
                    continue
                file_tabels_content = file_info.tabels_content
                cause_of_action = file_tabels_content['案由']
                if this_cause_of_action == cause_of_action:
                    table_father.display(self, "案由：正确。与《" + file_name + "》一致", "green")
                else:
                    table_father.display(self,
                                         "案由：与《" + file_name + "》不一致" + file_name + '中的案由为：' + str(
                                             cause_of_action),
                                         "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '案由：',
                                       '案由与《' + file_name + '》不一致' + file_name + '中的案由为：' + str(
                                           cause_of_action))

        # 针对案由需使用正则表达式的法律文书
        file_unregular_list = ['延长调查期限告知书', '先行登记保存证据处理通知书', '涉案物品返还清单']
        for file_unregular in file_unregular_list:
            if not tyh.file_exists(self.source_prifix, file_unregular):
                table_father.display(self, "文件缺失：《" + file_unregular + "》不存在", "red")
            else:
                temp_info = tyh.file_exists_open(self.source_prifix, file_unregular, DocxData)
                if temp_info is None:
                    table_father.display(self,
                                         "文件读取失败：" + file_unregular + "不存在，无法与《案件集体讨论记录》进行案由对比")
                    continue
                filetext = temp_info.text
                if this_cause_of_action not in filetext:
                    if file_unregular == file_unregular_list[1]:
                        tyh.addRemarkInDoc(self.mw, self.doc, '案由：',
                                           '案由与《延长调查期限告知书》中“案由”的记载不一致。')
                    elif file_unregular == file_unregular_list[2]:
                        tyh.addRemarkInDoc(self.mw, self.doc, '案由：',
                                           '案由与《先行登记保存证据处理通知书》中“案由”的记载不一致。')
                    else:
                        tyh.addRemarkInDoc(self.mw, self.doc, '案由：', '案由与《涉案物品返还清单》中“案由”的记载不一致。')

        # 针对案由在表格中的法律文书
        file_name_list = ['案件调查终结报告_']

        for file_name in file_name_list:
            if os.path.exists(self.source_prifix + file_name + ".docx") == 0:
                table_father.display(self, "文件缺失：《" + file_name + "》不存在", "red")
            else:
                file_info = DocxData(self.source_prifix + file_name + ".docx")
                file_tabels_content = file_info.tabels_content
                cause_of_action = file_tabels_content['案由']
                if this_cause_of_action == cause_of_action:
                    table_father.display(self, "案由：正确。与《" + file_name + "》一致", "green")
                else:
                    table_father.display(self,
                                         "案由：与《" + file_name + "》不一致" + file_name + '中的案由为：' + cause_of_action,
                                         "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, this_cause_of_action,
                                       '案由与《' + file_name + '》不一致' + file_name + '中的案由为：' + cause_of_action)

        # 针对案由需使用正则表达式的法律文书

    def check_time(self):
        """
        作用：立案日期：应当与《立案报告表》中的“负责人意见”的时间一致。
        """
        put_on_record_date = tyh.get_strtime(self.contract_tables_content['立案日期'])

        if not tyh.file_exists(self.source_prifix, "立案报告表"):
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            file_info = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
            file_tabels_content = file_info.tabels_content
            date_parttern = re.compile(r'.*(\d{4}年\d{1,2}月\d{1,2}日).*')
            date_by_principal = re.findall(date_parttern, file_tabels_content['负责人意见'])
            if date_by_principal == [''] or date_by_principal == []:
                table_father.display(self, "立案报告表_负责人意见_日期：应具体到XX年XX月XX日", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '立案日期', '立案报告表_负责人意见_日期 应具体到XX年XX月XX日')
            else:
                date_by_principal = tyh.get_strtime(date_by_principal[0])
                if not put_on_record_date or not date_by_principal:
                    table_father.display(self, "时间格式错误：立案日期或《立案报告表》中的“负责人意见”的时间格式有误",
                                         "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '立案日期',
                                       '立案日期或《立案报告表》中的“负责人意见”的时间格式有误,无法识别')
                elif put_on_record_date != date_by_principal:
                    table_father.display(self, "立案日期：与《立案报告表》中的“负责人意见”的时间不同", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '立案日期',
                                       '立案日期与《立案报告表》中的“负责人意见”的时间不同，《立案报告表》中的“负责人意见”的时间为：' + str(
                                           date_by_principal))
                else:
                    table_father.display(self, "立案日期：正确。与《立案报告表》中的“负责人意见”的时间相同", "green")

    def check_party_info(self):
        """
        作用：当事人：“姓名、性别、民族、证件类型及号码、住址”信息应当与《证据复制（提取）单》中当事人身份证明中记载的信息一致。
        """
        if '当事人' not in self.contract_tables_content or '性别' not in self.contract_tables_content or '民族' not in self.contract_tables_content or '证件类型号码' not in self.contract_tables_content:
            table_father.display(self,
                                 "文档格式错误，读取失败。请使用标准模板。请手动检查“姓名、性别、民族、证件类型及号码、住址”信息是否与《证据复制（提取）单》中当事人身份证明中记载的信息一致。",
                                 "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案件处理审批表',
                               '文档格式错误，读取失败。请使用标准模板。请手动检查“姓名、性别、民族、证件类型及号码、住址”信息是否与《证据复制（提取）单》中当事人身份证明中记载的信息一致。')
            return
        name = self.contract_tables_content['当事人'][1]
        gender = self.contract_tables_content['性别']
        nation = self.contract_tables_content['民族']
        certificate = self.contract_tables_content['证件类型号码']
        if "身份证" in certificate:
            is_IDcard = True
            temp = re.search(r"(\d+)", certificate)
            if temp:
                certificate = temp.group(1)
        else:
            is_IDcard = False
        address = self.contract_tables_content['住址']
        # tyh.addRemarkInDoc(self.mw, self.doc, '当事人', '姓名、性别、民族、证件类型及号码、住址”信息应当与《证据复制（提取）单》中当事人身份证明中记载的信息一致')
        comparelst = [name, gender, nation, address, certificate]
        file_list = [fn for fn in os.listdir(self.source_prifix) if ('~$' not in fn and (fn.endswith('.docx')))]
        # print(file_list)
        is_compareFile_exist = False  # 标志文件中是否存在至少一个《证据复制提取单》
        for f in file_list:
            if '证据复制提取单' in f:
                data = tyh.file_exists_open(self.source_prifix, f, DocxData)
                for i in data.tabels_content:
                    for j in i:
                        if "身份证" in j:
                            # print(f)
                            is_compareFile_exist = True
                            idInfoList = self.goGetIDCardInfo(f)
                            break
        if not is_compareFile_exist:
            table_father.display(self,
                                 "身份证信息缺失：文件夹中没有包含《证据复制提取单》或《证据复制提取单》不包含身份证信息！因此无法比较当事人信息！",
                                 "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '当事人',
                               '文件夹中没有包含《证据复制提取单》或《证据复制提取单》不包含身份证信息！因此无法比较当事人信息！')
        else:
            compareitem = ['姓名', '性别', '民族', '住址', '证件号码']
            remarkLoc = [comparelst[0], comparelst[1], comparelst[2], comparelst[3], comparelst[4]]
            if is_IDcard:
                i = 4
            else:
                table_father.display(self, "证件类型：不是身份证！后续跳过身份证号对比。", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '证件类型号码', "证件类型不是身份证！后续跳过身份证号对比。")
                i = 3
            while i >= 0:
                if not idInfoList[i] == comparelst[i]:
                    table_father.display(self, compareitem[
                        i] + "信息：与《证据复制提取单》的身份证图片中对应信息不相符！请人工审查", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, remarkLoc[i],
                                       compareitem[i] + "信息与《证据复制提取单》的身份证图片中对应信息不相符！请人工审查")
                i -= 1

    def goGetIDCardInfo(self, filename):
        pic_path = self.source_prifix + "picture/" + str(filename).strip('.docx') + "/word/media/"
        # print(pic_path)
        pic_list = [fn for fn in os.listdir(pic_path)]
        # print(pic_list)
        id_ocr = OCR_IDCard(pic_path + pic_list[0], pic_path + pic_list[1])
        id_name = id_ocr.getName()
        id_sex = id_ocr.getSex()
        id_nation = id_ocr.getNation()
        id_idnum = id_ocr.getIDnumber()
        id_address = id_ocr.getAddress()
        info = [id_name, id_sex, id_nation, id_address, id_idnum]
        return info

        # for item in iter(comparelst):
        #     if item == '/':
        #         comparelst.remove(item)

    def check_case_facts(self):
        """
        作用：案件事实：应当包括时间、地点、查获卷烟的品种、数量、金额、违法事实等。
        其中“时间”、“地点”应当与《立案报告表》中记载的一致，
        “品种”、“数量”应当与《物品清单》记载的一致。
        “金额”应当与《涉案烟草专卖品核价表》中的合计金额一致。
        违法事实应针对不同的案由设置不同的审查规则。
        """
        case_facts = self.contract_tables_content['案件事实']

        # 获取本表的信息
        time_parttern = re.compile(r'.*(\d{4}年\d{1,2}月\d{1,2}日).*')
        address_parttern = re.compile(r'')
        number_of_species_parttren = re.compile(r'共计(.*)个品种')
        money_parttern = re.compile(r'.*本案涉案金额为：(.*)元.*')

        self_time = re.findall(time_parttern, case_facts)[0]
        number_of_species = re.findall(number_of_species_parttren, case_facts)[0]
        money = re.findall(money_parttern, case_facts)
        # print(money)

        # 获取《立案报告表》的信息
        if not tyh.file_exists(self.source_prifix, "立案报告表"):
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        elif self_time != '':
            other_info_zero = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
            other_time = other_info_zero.tabels_content['案发时间']
            other_address = other_info_zero.tabels_content['案发地点']

            if self_time in other_time or self_time == other_time:
                table_father.display(self, "案发时间：正确。与《立案报告表》中记载的一致", "green")
            else:
                table_father.display(self, "案发时间：与《立案报告表》中记载的" + str(other_time) + "不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案件事实',
                                   '案发时间与《立案报告表》中记载的不一致,《立案报告表》中的时间为：' + str(other_time))
        else:
            if self_time == '' or self_time is None:
                table_father.display(self, "案发时间：不能为空", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案件事实', '案发时间不能为空')

        # 获取《涉案物品返回清单》的信息
        if not tyh.file_exists(self.source_prifix, "涉案物品返回清单"):
            table_father.display(self, "文件缺失：《涉案物品返回清单》不存在", "red")
        elif number_of_species != '':
            other_info_one = tyh.file_exists_open(self.source_prifix, "涉案物品返回清单", DocxData)
            other_info_one_len = len(other_info_one.tabels_content["涉案物品返还清单-品种"])
            # 首先确保物品种类相同
            if other_info_one_len == number_of_species:
                # 遍历涉案物品返还清单，判断其中的物品是否在案件处理审批表中
                for item in other_info_one.tabels_content["涉案物品返还清单-品种"]:
                    if item not in self.contract_tables_content['案件事实']:
                        table_father.display(self, "品种：与《物品清单》记载的不一致", "red")
                        tyh.addRemarkInDoc(self.mw, self.doc, '共计' + str(number_of_species) + '个品种',
                                           '品种与《涉案物品返还清单》记载的' + str(item) + '不一致')
            else:
                table_father.display(self, "品种的数量：与《物品清单》记载的不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '共计' + str(number_of_species) + '个品种',
                                   '品种的数量与《涉案物品返还清单》记载的不一致,《涉案物品返还清单》的数量为：' + str(
                                       other_info_one_len))
        else:
            table_father.display(self, "品种的数量：不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '共计' + str(number_of_species) + '个品种', '品种的数量不能为空')

        # 获取《涉案烟草专卖品核价表》的信息
        if not tyh.file_exists(self.source_prifix, "涉案烟草专卖品核价表"):
            table_father.display(self, "文件缺失：《涉案烟草专卖品核价表》不存在", "red")
        elif money != '' and money != []:
            other_info_two = tyh.file_exists_open(self.source_prifix, "涉案烟草专卖品核价表", DocxData)
            other_info_two_money = other_info_two.tabels_content['涉案烟草专卖品核价表-全部金额合计']
            if money[0] == other_info_two_money:
                table_father.display(self, '“金额”：正确。与《涉案烟草专卖品核价表》中的合计金额一致', 'green')
            else:
                table_father.display(self,
                                     '“金额”：与《涉案烟草专卖品核价表》中的合计金额不一致，《涉案烟草专卖品核价表》中的合计金额为：' + str(
                                         other_info_two_money),
                                     'red')
                tyh.addRemarkInDoc(self.mw, self.doc, '本案涉案金额为：',
                                   '“金额”与《涉案烟草专卖品核价表》中的合计金额不一致，《涉案烟草专卖品核价表》中的合计金额为：' + str(
                                       other_info_two_money))
        else:
            table_father.display(self, '涉案金额：未捕捉到涉案金额信息', 'red')
            tyh.addRemarkInDoc(self.mw, self.doc, '案件事实', '未捕捉到涉案金额信息')

    def check_basis_of_punishment(self):
        """
        作用：处罚依据：应针对不同的案由确定不同的审查规则。
        """
        table_father.display(self, "处罚依据：应针对不同的案由确定不同的审查规则")
        tyh.addRemarkInDoc(self.mw, self.doc, '处罚依据', '应针对不同的案由确定不同的审查规则')

    def check_opinions_of_the_undertaker(self):
        """
        作用：承办人意见：应当与《案件集体讨论记录》中的“结论性意见”基本一致，检索关键词：“没收”、“罚款”及数据，如不一致，应重点预警提示。
        “签名”应当在2人以上，
        “日期”应当在《案件集体讨论记录》的时间之后。
        """
        if not tyh.file_exists(self.source_prifix, "案件集体讨论记录"):
            table_father.display(self, "文件缺失：《案件集体讨论记录》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "案件集体讨论记录", DocxData)
            other_info_text = other_info.text

            # “日期”应当在《案件集体讨论记录》的时间之后。
            other_date_parttern = re.compile(r'与会人员签署意见并签名：.*(\d{4}年\d{1,2}月\d{1,2}日).*')
            self_date_parttern = re.compile(r'.*(\d{4}年\d{1,2}月\d{1,2}日).*')

            other_date = tyh.get_strtime(re.findall(other_date_parttern, other_info_text)[0])
            self_date = tyh.get_strtime(re.findall(self_date_parttern, self.contract_tables_content['承办人意见'])[0])

            time_differ = tyh.time_differ(self_date, other_date)
            if time_differ <= 0:
                table_father.display(self, "承办人意见日期：应当在《案件集体讨论记录》的时间之后", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '承办人\r意见', '日期应当在《案件集体讨论记录》的时间之后')
            else:
                table_father.display(self, "承办人意见日期：正确。在《案件集体讨论记录》的时间之后", "green")

            #
            tyh.addRemarkInDoc(self.mw, self.doc, '案件承办人员汇报案情及对案件的处理意见',
                               '应当与《案件集体讨论记录》中的“结论性意见”基本一致')
            tyh.addRemarkInDoc(self.mw, self.doc, '案件承办人员汇报案情及对案件的处理意见',
                               '请重点审查“没收”、“罚款”及数据')
            tyh.addRemarkInDoc(self.mw, self.doc, '签名', '“签名”应当在2人以上')

    def check_opinions_of_the_department(self):
        """
        作用：承办部门意见：检索关键词“同意”、“不同意”，必须有其中一个，
        如果检索出“不同意”，应当预警。签名完整，日期应在“承办人意见”的日期之后。
        """
        opinions_of_the_undertaker = self.contract_tables_content['承办人意见']
        opinions_of_the_department = self.contract_tables_content['承办部门意见']
        if '不同意' in opinions_of_the_department:
            table_father.display(self, "承办部门意见：为不同意！", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '预警！承办部门意见为不同意！')
        elif '同意' in opinions_of_the_department:
            table_father.display(self, "承办部门意见：正确。为同意", "green")
        else:
            table_father.display(self, "承办部门意见：不含有关键词“同意”、“不同意”二者之一", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '承办部门意见不含有关键词“同意”、“不同意”二者之一')

        self_date_parttern = re.compile(r'.*(\d{4}年\d{1,2}月\d{1,2}日).*')
        undertaker_time = tyh.get_strtime(re.findall(self_date_parttern, opinions_of_the_undertaker)[0])
        department_time = tyh.get_strtime(re.findall(self_date_parttern, opinions_of_the_department)[0])
        time_differ = tyh.time_differ(department_time, undertaker_time)
        if time_differ > 0:
            table_father.display(self, "承办部门意见日期：正确。在“承办人意见”的日期之后", "green")
        else:
            table_father.display(self, "承办部门意见日期：在“承办人意见”的日期之前", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '承办部门\r意见', '日期应在“承办人意见”的日期之后')

    def check_opinions_of_the_law_department(self):
        """
        作用：法制部门意见：检索关键词“同意”、“不同意”，必须有其中一个，
        如果检索出“不同意”，应当预警。签名完整，日期应在“承办人意见”的日期之后。
        """
        opinions_of_the_undertaker = self.contract_tables_content['承办人意见']
        opinions_of_the_law_department = self.contract_tables_content['法制部门意见']
        if '不同意' in opinions_of_the_law_department:
            table_father.display(self, "法制部门意见：为不同意！", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '法制部门意见', '预警！法制部门意见为不同意！')
        elif '同意' in opinions_of_the_law_department:
            table_father.display(self, "法制部门意见：正确。为同意", "green")
        else:
            table_father.display(self, "法制部门意见：不含有关键词“同意”、“不同意”二者之一", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '法制部门意见', '法制部门意见不含有关键词“同意”、“不同意”二者之一')

        self_date_parttern = re.compile(r'.*(\d{4}年\d{1,2}月\d{1,2}日).*')
        undertaker_time = tyh.get_strtime(re.findall(self_date_parttern, opinions_of_the_undertaker)[0])
        law_department_time = tyh.get_strtime(re.findall(self_date_parttern, opinions_of_the_law_department)[0])
        time_differ = tyh.time_differ(law_department_time, undertaker_time)
        if time_differ > 0:
            table_father.display(self, "法制部门意见日期：正确。在“承办人意见”的日期之后", "green")
        else:
            table_father.display(self, "法制部门意见日期：在“承办人意见”的日期之前", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '法制部门意见', '日期应在“承办人意见”的日期之后')

    def check_opinions_of_the_director(self):
        """
        作用：负责人意见：必须检索出“同意”或“不同意”，日期应当在法制部门意见日期之后，签名完整。
        """
        opinions_of_the_director = self.contract_tables_content['负责人意见']
        opinions_of_the_department = self.contract_tables_content['法制部门意见']
        if '不同意' in opinions_of_the_director:
            table_father.display(self, "负责人意见：为不同意！", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人 意见', '预警！负责人意见为不同意！')
        elif '同意' in opinions_of_the_director:
            table_father.display(self, "负责人意见：正确。为同意", "green")
        else:
            table_father.display(self, "负责人意见：不含有关键词“同意”、“不同意”二者之一", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人 意见', '负责人意见不含有关键词“同意”、“不同意”二者之一')

        self_date_parttern = re.compile(r'.*(\d{4}年\d{1,2}月\d{1,2}日).*')
        director_time = tyh.get_strtime(re.findall(self_date_parttern, opinions_of_the_director)[0])
        department_time = tyh.get_strtime(re.findall(self_date_parttern, opinions_of_the_department)[0])
        time_differ = tyh.time_differ(director_time, department_time)
        if time_differ > 0:
            table_father.display(self, "负责人意见日期：正确。在“法制部门意见”的日期之后", "green")
        else:
            table_father.display(self, "负责人意见日期：在“法制部门意见”的日期之前", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '负责人意见', '日期应在“法制部门意见”的日期之后')

    def check_additional_function_one_ner(self):
        """
        date:2022.2.20
        function:
        1.日期应与检查（勘验）时间的开始时间一致
        2.地点应与检查（勘验）地点一致
        3.事件事实文本中是否含有执法人员要素
        """
        # target_file_name_index = self.is_target_file_exit(self.source_prifix, self.target_file_name_list[4])
        # if target_file_name_index == -1:
        #     table_father.display(self, "× 检查（勘验）笔录.docx不存在", "red")
        if not tyh.file_exists(self.source_prifix, "检查（勘验）笔录"):
            table_father.display(self, "文件缺失：《检查（勘验）笔录》不存在", "red")
        else:
            # 获取《检查（勘验）笔录》的时间与地点部分文字
            other_context = tyh.file_exists_open(self.source_prifix, "检查（勘验）笔录", DocxData).text
            other_time = re.findall("检查（勘验）时间：(.*?)\n", other_context)[0]
            other_address = re.findall("检查（勘验）地点：(.*?)\n", other_context)[0]
            # 使用NER获取案件处理审批表的时间，取第一个出现
            cognitio = self.contract_tables_content['案件事实']
            this_time_for_check = self.entityrecognition.get_identity_with_tag(cognitio, "TIME")[0]
            # print(this_time_for_check)

            if this_time_for_check in other_time:
                table_father.display(self, "案件事实日期：正确。与检查（勘验）时间的开始时间一致", "green")
            else:
                table_father.display(self, '日期不符：案件事实中的日期为' + str(
                    this_time_for_check) + ',与检查（勘验）时间的开始时间不一致', "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案件事实',
                                   '案件事实中的日期为' + str(
                                       this_time_for_check) + ',与检查（勘验）时间的开始时间不一致')

            if other_address in cognitio:
                table_father.display(self, "案件事实：正确。中的地点与检查（勘验）地点一致", "green")
            else:
                table_father.display(self, "案件事实地点：与检查（勘验）地点不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案件事实', '案件事实中的地点与检查（勘验）地点不一致')

            if "执法人员" not in cognitio:
                table_father.display(self, "案件事实：不包含执法人员", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案件事实', '案件事实中不包含执法人员')

    def check_additional_funtion_two_ner(self):
        """
        date:2022.2.20
        function:
        案件事实的无规则文本中是否含有烟草专卖零售许可证、卷烟数目、法律条款等要素
        """
        cognitio = self.contract_tables_content['案件事实']
        if "烟草专卖零售许可证" not in cognitio:
            table_father.display(self, "案件事实：不包含烟草专卖零售许可证", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案件事实', '案件事实中不包含烟草专卖零售许可证')
        if "卷烟数目" not in cognitio:
            table_father.display(self, "案件事实：不包含卷烟数目", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案件事实', '案件事实中不包含卷烟数目')
        if "法律条款" not in cognitio:
            table_father.display(self, "案件事实：不包含法律条款", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案件事实', '案件事实中不包含法律条款')

    def get_anyou(self):
        """
        date:2022.6.27
        function: 
        返回return_info_one = [{'案由':'未在当地烟草专卖批发企业进货', '处罚结果':['处以xxx。','予以xxx。']}, {'案由':'未在当地烟草专卖批发企业进货', '处罚结果':['处以xxx。','予以xxx。']}]
           return_info_two = {'从轻':True or False,'存在从轻证据':['xxxxxx','xxxx'],'从重':True or False}
        返回格式：return_info_one,return_info_two
        流程：分别获取案由与处罚结果。
             判断案由字符串的某一比例（如80%）是否与某一条处罚结果匹配。若匹配，处理数据准备上传。若不匹配，提示该处罚结果未与已有案由存在匹配关系，人工审查。
             修改从轻从重的bool值，将打包好的字典塞入return_info_one
        """
        # interface_info_two = ['《' + x for x in
        #                       list(filter(_not_empty, self.contract_tables_content['处罚依据'].split("《")))]

        if '当事人' not in self.contract_tables_content or '性别' not in self.contract_tables_content or '民族' not in self.contract_tables_content or '证件类型号码' not in self.contract_tables_content:
            table_father.display(self,
                                 "文档格式错误",
                                 "文档格式错误。姓名、性别、民族、证件类型及号码、住址读取失败。请使用标准模板。请手动检查处罚结果是否与已有案由存在匹配关系",
                                 "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '案件处理审批表',
                               '文档格式错误。姓名、性别、民族、证件类型及号码、住址读取失败。请使用标准模板。请手动检查处罚结果是否与已有案由存在匹配关系')
            return [], {'从轻': False, '存在从轻证据': [], '从重': False}
        # 涉及的变量
        ratio_ = 0.8  # 案由子字符串占比（默认0.8）
        brief = re.split(r'[,，;；、]', self.contract_tables_content['案由'])
        # undertakers_opinion = re.findall(r"(?<!可|并)[处|予]以.*?[。|;|；]", self.contract_tables_content['承办人意见'])
        undertakers_opinion = []
        split_result = re.split(r'[;；。]', self.contract_tables_content['承办人意见'])
        for item in split_result:
            if '没收' in item or '销毁' in item or '罚款' in item:
                undertakers_opinion.append(item)

        is_light = False  # 默认从轻处罚为Flase
        is_heavy = False  # 默认从重处罚为Flase

        nonage = False  # 默认成年

        evidence_library = ['违法行为人已满十四周岁不满十八周岁', '主动中止违法行为，且危害后果轻微',
                            '主动消除或者减轻违法行为危害后果',
                            '在共同违法行为中起次要或辅助作用', '受他人胁迫或诱骗实施违法行为',
                            '主动供述行政机关尚未掌握的违法行为',
                            '配合烟草专卖局查处违法行为有立功表现',
                            '法律、法规、规章规定应当从轻或者减轻行政处罚的其他情形']
        light_evidence = []  # 从轻证据

        return_info_one = []  # 返回信息一
        return_info_two = {'从轻': False, '存在从轻证据': [], '从重': False}  # 返回信息二

        # 遍历案由，并获取每个案由的子字符串，与处罚结果进行比较
        for anyou in brief:
            append_item = {'案由': anyou, '处罚结果': []}
            for substring in self.get_substring_by_ratio(anyou, ratio=ratio_):
                for chufajieguo in undertakers_opinion:

                    # 若处罚结果长度小于子字符串，跳过
                    if len(chufajieguo) < len(substring):
                        pass
                    # 否则，判断子字符串是否在某一处罚结果中。
                    # 若在，匹配结果。不在，跳过。
                    elif substring in chufajieguo:
                        append_item['处罚结果'].append(chufajieguo)
                        undertakers_opinion.remove(chufajieguo)
                    else:
                        continue
            if append_item['处罚结果'] is not None:
                return_info_one.append(append_item)

        # 对未被匹配的处罚结果进行提示
        if undertakers_opinion is not None:
            for item in undertakers_opinion:
                tyh.addRemarkInDoc(self.mw, self.doc, item, '该处罚结果未被已有案由匹配，请人工审查')

        # -----return_info_one above----------------------return_info_two below-------------

        # 身份证号码正则表达式匹配
        idcard_string = self.contract_tables_content['证件类型号码']
        idcard_number = re.search(r'([1-9]\d{5}[12]\d{3}(0[1-9]|1[012])(0[1-9]|[12][0-9]|3[01])\d{3}[0-9xX])',
                                  idcard_string, re.S)
        # 判断当事人是否未成年，若存在未成年情况，则从轻处罚,且增加从轻证据
        if idcard_number is None:
            tyh.addRemarkInDoc(self.mw, self.doc, '证件类型号码', '未识别到有效身份证号码')
        else:
            idcard_processor = IdCard_Information_Processor(idcard_number.group())
            age = idcard_processor.get_age()
            if 14 <= age < 18:
                is_light = True
                return_info_two['从轻'] = True
                light_evidence.append('违法行为人已满十四周岁不满十八周岁')
                nonage = True
                tyh.addRemarkInDoc(self.mw, self.doc, '证件类型号码', '当事人年龄为' + str(age) + '，应从轻处罚')
        # 若当事人已成年，执行其它判定
        if not nonage:
            if '从轻' in self.contract_tables_content['承办人意见']:
                is_light = True
                return_info_two['从轻'] = True

        if '从重' in self.contract_tables_content['承办人意见']:
            is_heavy = True
            return_info_two['从重'] = True

        for item in evidence_library:
            if item in self.contract_tables_content['承办人意见']:
                return_info_two['存在从轻证据'].append(item)

        return return_info_one, return_info_two

        # 提示处罚依据为空
        # if len(interface_info_two) == 0:
        #     table_father.display(self, "× 处罚依据为空", "red")
        #     tyh.addRemarkInDoc(self.mw, self.doc, '处罚依据', '处罚依据为空')
        # # 提示处罚依据与处罚结果数目不对应
        # elif len(interface_info_two) != len(interface_info_three):
        #     table_father.display(self, "× 处罚依据与处罚结果数目不对应", "red")
        #     tyh.addRemarkInDoc(self.mw, self.doc, '处罚依据', '处罚依据与处罚结果数目不对应')
        # else:
        #     for i, item in enumerate(interface_info_two):
        #         result_of_handling = Discretionary_Power(
        #             [interface_info_one, interface_info_two[i], interface_info_three[i]], is_light)
        #         # result_of_handling为一个长度为2的list。第一位为数字，1代表打标注的位置为【处罚依据】，2代表打标注的位置为【承办人意见】
        #         # 第二位则为将被标注的string
        #         table_father.display(self, result_of_handling[1], "red")
        #         if result_of_handling[0] == 1:
        #             # tyh.addRemarkInDoc(self.mw, self.doc, '处罚依据', result_of_handling[1])
        #             tyh.addRemarkInDoc(self.mw, self.doc, interface_info_two[i], result_of_handling[1])
        #         else:
        #             tyh.addRemarkInDoc(self.mw, self.doc, '承办人\r意见', result_of_handling[1])

        # 加一个从轻
        # print(interface_info_one)
        # print(interface_info_two)
        # print(interface_info_three)
        # return [interface_info_one, interface_info_two, interface_info_three]

    def get_substring_by_ratio(self, input_string, ratio=0.8):
        substring_len = math.floor(len(input_string) * ratio)
        results = []
        for x in range(len(input_string) - substring_len + 1):
            results.append(input_string[x:x + substring_len])
        return results

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
        self.doc = self.mw.Documents.Open(self.my_prefix + file_name_real)
        data = DocxData(file_path=contract_file_path)
        self.contract_text = data.text
        self.contract_tables_content = data.tabels_content
        for func in self.all_to_check:
            try:
                eval(func)
            except Exception as e:
                table_father.display(self,
                                     "文档格式有误，请主观审查下列功能：" + function_description_dict[str(func)[5:-2]],
                                     "red")
                table_father.display(self, "文档存在格式错误，函数失效：" + func + ' 遇到错误:' + str(e.args))
        self.doc.Save()
        self.doc.Close()

        self.mw.Quit()
        self.mw.Quit()
        print("《案件处理审批表》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\twj\\Desktop\\zzz\\10\\"
    # my_prefix = "C:\\Users\\Xie\\Desktop\\案卷组_成烟立测试文书 - 副本\\2021184250_成烟立[2021]第2号\\"
    dit_list = os.listdir(my_prefix)
    if "案件处理审批表_.docx" in dit_list:
        ioc = Table34(my_prefix, my_prefix)
        contract_file_path = my_prefix + "案件处理审批表_.docx"
        ioc.check(contract_file_path, "案件处理审批表_.docx")
    # my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36 - 副本\\"
    # dit_list = os.listdir(my_prefix)
    # if "案件处理审批表_.docx" in dit_list:
    #     ioc = Table34(my_prefix, my_prefix)
    #     contract_file_path = my_prefix + "案件处理审批表_.docx"
    #     ioc.check(contract_file_path, "案件处理审批表_.docx")
