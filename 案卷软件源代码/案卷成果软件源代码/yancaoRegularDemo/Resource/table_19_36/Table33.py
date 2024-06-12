import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
import os
import win32com
from win32com.client import Dispatch
from yancaoRegularDemo.Resource.ReadFile import DocxData
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh

function_description_dict = {
    'check_cause_of_action': '案由应当与《案件调查终结报告》之后的全部法律文书中“案由”的记载一致。',
    'check_serial_number': '案件编号应当与《立案报告表》中的编号一致。',
    'check_time': '时间应当在《案件处理审批表》记载的时间之前。',
    'check_site': '地点不为空',
    'check_compere': '主持人不为空',
    'check_compere_duty': '主持人职务：应当为“局长”或“副局长”',
    'check_present': '出席人员姓名及职务：姓名个数应当为单数。',
    'check_handling_suggestion': '“处理意见”一般应当与《案件调查终结报告》中的“处理意见”一致。',
    'check_discuss_record': '各出席人员，均应当有发言内容',
    'check_conclusion': '结论性意见：应针对不同的案由设置不同的审查规则。',
    'check_sign': '结论性意见：应针对不同的案由设置不同的审查规则。',
}


# 案件集体讨论记录
class Table33(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.my_prefix = my_prefix
        self.source_prifix = source_prifix  # 2021-08-07版本新增
        self.mw = win32com.client.Dispatch("Word.Application")
        self.mw.Visible = 0
        self.mw.DisplayAlerts = 0
        self.contract_text = None
        self.contract_tables_content = None

        self.all_to_check = [
            "self.check_cause_of_action()",
            "self.check_serial_number()",
            "self.check_time()",
            "self.check_site()",
            "self.check_compere()",
            "self.check_compere_duty()",
            "self.check_present()",
            "self.check_handling_suggestion()",
            "self.check_discuss_record()",
            "self.check_conclusion()",
            "self.check_sign()"
        ]

    def check_cause_of_action(self):
        """
        作用：案由：应当与《案件调查终结报告》之后的全部法律文书中“案由”的记载一致。
        """
        this_cause_of_action_parttern = re.compile(r'案由：(.*)\n*案件编号：')
        this_cause_of_action = re.findall(this_cause_of_action_parttern, self.contract_text)[0]

        # 针对案由在表格中的法律文书
        file_with_table_list = ['案件调查终结报告', '延长调查终结审批表']

        for file_name in file_with_table_list:
            if not tyh.file_exists(self.source_prifix, file_name):
                table_father.display(self, "文件缺失：" + "《" + file_name + "》不存在", "red")
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
                                         "案由：与《" + file_name + "》不一致，" + file_name + '中的案由为：' + cause_of_action,
                                         "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '案由：',
                                       "案由与《" + file_name + "》不一致，" + file_name + '中的案由为：' + cause_of_action)

        # 针对案由需使用正则表达式的法律文书
        file_unregular_list = ['延长调查期限告知书', '先行登记保存证据处理通知书', '涉案物品返还清单']
        for file_unregular in file_unregular_list:
            if not tyh.file_exists(self.source_prifix, file_unregular):
                table_father.display(self, "文件缺失：《" + file_unregular + "》不存在", "red")
            else:
                temp_info = tyh.file_exists_open(self.source_prifix, file_unregular, DocxData)
                if temp_info is None:
                    table_father.display(self,
                                         "文件缺失：" + file_unregular + "不存在，无法与《案件集体讨论记录》进行案由对比")
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

    def check_serial_number(self):
        """
        作用：案件编号：应当与《立案报告表》中的编号一致。
        """
        if not tyh.file_exists(self.source_prifix, "立案报告表"):
            table_father.display(self, "文件缺失：《立案报告表》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "立案报告表", DocxData)
            other_text = other_info.text
            this_parttern = re.compile(r'案件编号：(.*)[\s]*时间：')
            this_serial_number = re.findall(this_parttern, self.contract_text)[0]
            if this_serial_number in other_text:
                table_father.display(self, "案件编号：正确。与《'立案报告表》中编号一致", "green")
            else:
                table_father.display(self, "案件编号：与《立案报告表》中编号不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案件编号：', '案件编号与《立案报告表》中的编号不一致')

    def check_time(self):
        """
        作用：时间：应当在《案件处理审批表》记载的时间之前。（讨论先告知或是先讨论）
        """
        if not tyh.file_exists(self.source_prifix, "案件处理审批表"):
            table_father.display(self, "文件缺失：《案件处理审批表》不存在", "red")
        else:
            file_info = tyh.file_exists_open(self.source_prifix, "案件处理审批表", DocxData)
            file_tabels_content = file_info.tabels_content
            table_father.display(self, "时间：提示。时间应当在《案件处理审批表》记载的时间之前。请主观审查。", "red")
        tyh.addRemarkInDoc(self.mw, self.doc, '时间', "时间应当在《案件处理审批表》记载的时间之前。请主观审查。")

    def check_site(self):
        """
        作用：地点：不为空，非审查要点。
        """
        site_parttern = re.compile(r'地点：(.*)\n*主持人：')
        site = re.findall(site_parttern, self.contract_text)[0]
        site = str(site).rstrip()
        if site == '' or site is None:
            table_father.display(self, "地点：为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '地点：', '地点为空')
        else:
            table_father.display(self, "地点：正确。不为空", "green")

    def check_compere(self):
        """
        作用：主持人：不为空。
        """
        compere_parttern = re.compile(r'主持人：(.*)\n*职务：')
        compere = re.findall(compere_parttern, self.contract_text)[0]
        compere = str(compere).rstrip()
        if compere == '' or compere is None:
            table_father.display(self, "主持人：为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '主持人：', '主持人:为空')
        else:
            table_father.display(self, "主持人：正确。不为空", "green")

    def check_compere_duty(self):
        """
        作用：主持人职务：应当为“局长”或“副局长”
        """
        compere_duty_list = ['局长', '副局长']
        compere_duty_parttern = re.compile(r'职务：(.*)\n*出席人员姓名及职务：')
        compere_duty = re.findall(compere_duty_parttern, self.contract_text)[0]
        compere_duty = str(compere_duty).lstrip()
        if compere_duty in compere_duty_list:
            table_father.display(self, "主持人职务：正确。为“局长”或“副局长”", "green")
        else:
            table_father.display(self, "× 主持人职务：不为“局长”或“副局长”之一", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '职务：', '主持人职务：不为“局长”或“副局长”')

    def check_present(self):
        """
        作用：出席人员姓名及职务：姓名个数应当为单数。
        """
        all_name_parttern = re.compile(r'出席人员姓名及职务：(.*)记录人：', re.S)
        all_name = re.findall(all_name_parttern, self.contract_text)[0]
        name_list = str(all_name).replace('，', '、').replace(",", '、').replace("\n", "").split('、')
        if len(name_list) % 2 == 1:
            table_father.display(self, "姓名个数：正确。为单数", "green")
        else:
            table_father.display(self, "姓名个数：为双数。应当为单数。", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '出席人员姓名及职务：', '姓名个数应当为单数')

    def check_handling_suggestion(self):
        """
        作用：应当有“案件承办人员汇报案情及对案件的处理意见”的记载。
            关键词“汇报案情”、“处理意见”。“处理意见”一般应当与《案件调查终结报告》中的“处理意见”一致。
        """
        if not tyh.file_exists(self.source_prifix, "案件调查终结报告"):
            table_father.display(self, "文件缺失：《案件调查终结报告》不存在", "red")
        else:
            other_info = tyh.file_exists_open(self.source_prifix, "案件调查终结报告", DocxData)
            other_handling_suggestion = str(other_info.tabels_content['处理意见'])
            other_handling_suggestion = other_handling_suggestion[0:other_handling_suggestion.find('签名')]
            if other_handling_suggestion in self.contract_text:
                table_father.display(self, "“处理意见”：正确。与《案件调查终结报告》中的“处理意见”一致", "green")
            else:
                table_father.display(self, "“处理意见”：与《案件调查终结报告》中的“处理意见”不一致", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '案件承办人员汇报案情及对案件的处理意见',
                                   '“处理意见”与《案件调查终结报告》中的“处理意见”不一致')

    def check_discuss_record(self):
        """
        作用：“讨论记录”，将第7项中出席人员姓名作为关键词，每个姓名下，均应当有发言内容。
        """
        # 将第7项中出席人员姓名作为关键词
        all_name_parttern = re.compile(r'出席人员姓名及职务：(.*)记录人：', re.S)
        all_name = re.findall(all_name_parttern, self.contract_text)[0]
        name_list = str(all_name).replace('，', '、').replace(",", '、').replace("\n", "").split('、')
        new_name_list = []
        # 去除第7项中出席人员姓名后的职称
        for name in name_list:
            new_name_list.append(name.split('（')[0])

        # 获取全部发言记录(包含\n)
        all_record_parttern = re.compile(r'讨论记录：(.*)结论性意见：', re.S)
        all_record = re.findall(all_record_parttern, self.contract_text)[0]

        # 遍历姓名
        for name in new_name_list:
            singal_record_parttern_str = name + '[：:]*\n(.*)\n'
            singal_record_parttern = re.compile(singal_record_parttern_str)
            single_record = re.findall(singal_record_parttern, all_record)[0]
            if single_record == '' or single_record is None:
                table_father.display(self, "发言记录：" + "未查询到" + name + "的发言记录", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '讨论记录：', "未查询到" + name + "的发言记录")

    def check_conclusion(self):
        """
        作用：结论性意见：应针对不同的案由设置不同的审查规则。
            比如：未在当地烟草专卖企业进货案件，检索关键词“罚款”及比例，比例值应当在10%以下 。
        """

    def check_sign(self):
        """
        作用：与会人员签署意见并签名：签名应当与第7项中出席人员姓名对应，
            签署意见应当为“同意”“不同意”“保留”或者“弃权”中的一种。
        """
        table_father.display(self,
                             "签名：提示。签名应当与第7项中出席人员姓名对应，签署意见应当为“同意”“不同意”“保留”或者“弃权”中的一种。请主观审查。",
                             "red")
        tyh.addRemarkInDoc(self.mw, self.doc, '与会人员签署意见并签名',
                           "签名应当与第7项中出席人员姓名对应；签署意见应当为“同意”“不同意”“保留”或者“弃权”中的一种；请主观审查。")

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
            # eval(func)
        self.doc.Save()
        self.doc.Close()

        self.mw.Quit()
        # self.mw.Quit()
        print("《案件集体讨论记录》审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Xie\\Desktop\\Table19-36\\"
    list = os.listdir(my_prefix)
    if "案件集体讨论记录_.docx" in list:
        ioc = Table33(my_prefix, my_prefix)
        contract_file_path = my_prefix + "案件集体讨论记录_.docx"
        ioc.check(contract_file_path, '案件集体讨论记录_.docx')
