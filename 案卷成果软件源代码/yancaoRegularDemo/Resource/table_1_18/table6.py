from yancaoRegularDemo.Resource.tools.EntityRecognition import EntityRecognition
from yancaoRegularDemo.Resource.tools.tangyuhao_readfile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *

function_description_dict = {
    'timeRight': '涉及举报的，检查时间在举报时间之后。',
    'peopleRight': '检查地点、住址、证件类型及号码与《立案报告表》保持一致。性别、联系电话、烟草专卖许可证号码、现场负责人主观审查。',
    'elseRight': '每一页被检查人名字及签名时间、执法人员名字及签名时间保持一致，且签名时间注意与勘验时间保持一致。若被检查人拒签，预警提示。',
    'xianchangqingkuang_gaozhishixiang_case1': '现场情况部分：必备要素：应包括检查时间、检查地点、执法人员名字及执法号码（涉及多部门联合检查，比如公安、邮政、市监的，主观审查），'
                                               '出示检查证件。主观审查要素：涉案卷烟品牌、规格、数量，违法条款，是否经过领导批准。',
}


class table6(table_father):
    def __init__(self, my_prefix, source_prifix):
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix
        self.entityrecognition = EntityRecognition()

        self.all_to_check = [
            "self.timeRight()",
            "self.peopleRight()",
            "self.elseRight()",
            "self.xianchangqingkuang_gaozhishixiang_case1()"

        ]

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
        self.mw = win32com.client.Dispatch("Word.Application")
        self.doc = self.mw.Documents.Open(self.my_prefix + file_name_real)
        data = file_1(file_path=contract_file_path)
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
        # self.mw.Quit()
        print(file_name_real + "审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result

    def timeRight(self):
        text0 = self.contract_text
        time = re.findall("检查（勘验）时间：(.*?)\n", text0)[0]
        pattern = r'(\d{4}年\d{1,2}月\d{1,2}日\d{1,2}时\d{1,2}分至\d{4}年\d{1,2}月\d{1,2}日\d{1,2}时\d{1,2}分)'
        if re.match(pattern, time) == None:
            table_father.display(self, "检查（勘验）时间：检查时间错误", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '检查（勘验）时间：', '检查时间错误')
            return
        if "至" in time:
            time = re.findall("(.*)至.*?", time)[0]
            time = time.replace("年", "-").replace("月", "-").replace("日", "-").replace("时", "-").replace("分",
                                                                                                            "").replace(
                "/",
                "-").strip()

        # if os.path.exists(self.my_prefix + "举报记录表_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "举报记录表"):
            data1 = tyh.file_exists_open(self.source_prifix, "举报记录表", file_1)
            time0 = data1.tabels_content['举报时间']
            time0 = re.findall("(\d{4}年\d{1,2}月\d{1,2}日\d{1,2}时\d{1,2}分)", time0)
            if time0 != [] or time0 != ['']:
                time0 = time0[0]
                time0 = time0.replace("年", "-").replace("月", "-").replace("日", "-").replace("时", "-"). \
                    replace("分", "").replace("/", "-").strip()

                time0 = re.findall(r"\d+\.?\d*", time0)

                time = re.findall(r"\d+\.?\d*", time)

                if int(time[0]) > int(time0[0]) or int(time[1]) > int(time0[1]) or int(time[2]) > int(
                        time0[2]) or int(time[3]) > int(time0[3]) or \
                        int(time[3]) * 60 + \
                        int(time[4]) - int(time0[3]) * 60 - int(time0[4]) >= 0:
                    table_father.display(self, "检查（勘验）时间：检查时间在举报时间（" + str(time0) + "）之后", "green")

                else:
                    table_father.display(self,
                                         "检查（勘验）时间：检查时间（" + str(time) + "）在举报时间（" + str(time0) + "）之前",
                                         "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '检查（勘验）时间：',
                                       "检查时间（" + str(time) + "）在举报时间（" + str(time0) + "）之前")

    def peopleRight(self):
        check_place = ''
        pattern = r'.*检查（勘验）地点：(.*)'
        text0 = re.findall(pattern, self.contract_text)
        if text0 == [] or text0[0].strip() == "" or text0[0].replace(" ", "") == '/':
            table_father.display(self, "检查（勘验）地点：检查（勘验）地点不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '检查（勘验）地点：', '检查（勘验）地点不能为空')
        else:
            check_place = text0[0]

        check_id = ''
        pattern = r'.*证件类型及号码：(.*)'
        text0 = re.findall(pattern, self.contract_text)
        if text0 == [] or text0 == [""] or text0 == ['/']:
            table_father.display(self, "证件类型及号码：证件类型及号码不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '证件类型及号码：', '证件类型及号码不能为空')
        else:
            check_id = text0[0]

        check_zhuzhi = ''
        pattern = r'.*住址：(.*)'
        text0 = re.findall(pattern, self.contract_text)
        if text0 == [] or text0 == [""] or text0 == ['/']:
            table_father.display(self, "住址：住址不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '住址：', '住址不能为空')
        else:
            check_zhuzhi = text0[0]

        # if os.path.exists(self.my_prefix + "立案报告表_.docx") == 1:
        if tyh.file_exists(self.source_prifix, "立案报告表"):
            data1 = tyh.file_exists_open(self.source_prifix, "立案报告表", file_1)

            place = data1.tabels_content['案发地点']
            id = data1.tabels_content['证件类型及号码']
            zhuzhi = data1.tabels_content['地址']
            if check_place.replace(" ", "") != '' and place.replace(" ", "") != '':
                if check_place.replace(" ", "") != place.replace(" ", ""):
                    table_father.display(self, "检查（勘验）地点：检查（勘验）地点（" + str(
                        check_place) + "）与立案报告表“案发地点”（" + str(place) + "）不一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '检查（勘验）地点：',
                                       "检查（勘验）地点（" + str(check_place) + "）与立案报告表“案发地点”（" + str(
                                           place) + "）不一致")

            if check_id.replace(" ", "") != '' and id.replace(" ", "") != '':
                if check_id.replace(" ", "") != id.replace(" ", ""):
                    table_father.display(self, "证件类型及号码：检查（勘验）笔录“证件类型及号码”（" + str(
                        check_id) + "）与立案报告表“证件类型及号码”（" + str(id) + "）不一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '证件类型及号码：', "检查（勘验）笔录“证件类型及号码”（" + str(
                        check_id) + "）与立案报告表“证件类型及号码”（" + str(id) + "）不一致")

            if check_zhuzhi.replace(" ", "") != '' and zhuzhi.replace(" ", "") != '':
                if '：' in zhuzhi:
                    zhuzhi = zhuzhi.split('：')[1]
                if '：' in check_zhuzhi:
                    check_zhuzhi = check_zhuzhi.split('：')[1]
                if check_place != place:
                    table_father.display(self, "住址：检查（勘验）笔录“住址”（" + str(
                        check_place) + "）与立案报告表“案发地点”（" + str(place) + "）不一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '住址：',
                                       "检查（勘验）笔录“住址”（" + str(check_place) + "）与立案报告表“案发地点”（" + str(
                                           place) + "）不一致")

        tyh.addRemarkInDoc(self.mw, self.doc, '性别：', '主观审查')
        tyh.addRemarkInDoc(self.mw, self.doc, '联系电话：', '主观审查')
        tyh.addRemarkInDoc(self.mw, self.doc, '烟草专卖许可证号码：', '主观审查')
        tyh.addRemarkInDoc(self.mw, self.doc, '现场负责人：', '主观审查')
        table_father.display(self, "主观审查：性别、联系电话、烟草专卖许可证号码、现场负责人主观审查。", "red")

        # pattern=r'.*被检查（勘验）人名称：(.*)法定代表人*|\n'
        text0 = ""
        regex_list = [
            ".*被检查（勘验）人名称：(.*)法定代表人*",
            ".*被检查（勘验）人名称：(.*)\n"
        ]
        for regex in regex_list:
            t = re.findall(regex, self.contract_text)
            if t:
                text0 = t
                break
        # text=re.findall(pattern,self.contract_text)
        if text0 == [] or text0[0].replace(" ", "") == "" or text0[0].replace(" ", "") == '/':
            table_father.display(self, "被检查（勘验）人名称：被检查（勘验）人名称不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '被检查（勘验）人名称：', '被检查（勘验）人名称不能为空')

        pattern = r'.*法定代表人（负责人）：(.*)'
        text0 = re.findall(pattern, self.contract_text)
        if text0 == [] or text0[0].replace(" ", "") == "" or text0[0].replace(" ", "") == '/':
            table_father.display(self, "法定代表人（负责人）：法定代表人（负责人）不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '法定代表人（负责人）：', '法定代表人（负责人）不能为空')

        text0 = ""
        regex_list = [
            ".*被检查（勘验）人姓名：(.*)性别*",
            ".*被检查（勘验）人姓名：(.*)\n"
        ]
        for regex in regex_list:
            t = re.findall(regex, self.contract_text)
            if t:
                text0 = t
                break
        # pattern=r'.*被检查（勘验）人姓名：(.*)'
        # text=re.findall(pattern,self.contract_text)
        if text0 == [] or text0[0].replace(" ", "") == "" or text0[0].replace(" ", "") == '/':
            table_father.display(self, "被检查（勘验）人姓名：被检查（勘验）人姓名不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '被检查（勘验）人姓名：', '被检查（勘验）人姓名不能为空')

        pattern = r'.*与被检查（勘验）人关系：(.*)'
        text0 = re.findall(pattern, self.contract_text)
        if text0 == [] or text0[0].replace(" ", "") == "" or text0[0].replace(" ", "") == '/':
            table_father.display(self, "与被检查（勘验）人关系：与被检查（勘验）人关系不能为空", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '与被检查（勘验）人关系：', '与被检查（勘验）人关系不能为空')

        tyh.addRemarkInDoc(self.mw, self.doc, '现场情况：',
                           '不为空，必备要素：应包括检查时间、检查地点、执法人员名字及执法号码（涉及多部门联合检查，比如公安、邮政、市监的，主观审查），出示检查证件。主观审查要素：涉案卷烟品牌、规格、数量，违法条款，是否经过领导批准。')

    def elseRight(self):
        pattern = r'.*被检查（勘验）人或现场负责人（签名）：(.*)\n'
        list1 = re.findall(pattern, self.contract_text)
        if tyh.allSame_noSpace(list1) == False:
            table_father.display(self, "被检查（勘验）人或现场负责人（签名）：被检查（勘验）人或现场负责人（签名）或日期（" + str(
                list1) + "）不一致", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '被检查（勘验）人或现场负责人（签名）：',
                               "被检查（勘验）人或现场负责人（签名）或日期（" + str(list1) + "）不一致")
        else:
            table_father.display(self, "被检查（勘验）人或现场负责人（签名）：被检查（勘验）人或现场负责人（签名）和日期一致",
                                 "green")
            tyh.addRemarkInDoc(self.mw, self.doc, '被检查（勘验）人或现场负责人（签名）：', '主观审查是否为空')

        pattern = r'.*见证人（签名）：(.*)\n'
        list2 = re.findall(pattern, self.contract_text)
        if tyh.allSame_noSpace(list2) == False:
            table_father.display(self, "见证人（签名）：见证人（签名）或日期（" + str(list2) + "）不一致", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '见证人（签名）：', "见证人（签名）或日期（" + str(list2) + "）不一致",
                               "red")
        else:
            table_father.display(self, "见证人（签名）：见证人（签名）和日期一致", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, '见证人（签名）：', '主观审查是否为空')

        pattern = r'.*检查（勘验）人（签名）：(.*)\n'
        list3 = re.findall(pattern, self.contract_text)
        if tyh.allSame_noSpace(list3) == False:
            table_father.display(self, "检查（勘验）人（签名）：检查（勘验）人（签名）：或日期（" + str(list3) + "）不一致", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '检查（勘验）人（签名）：',
                               "检查（勘验）人（签名）：或日期（" + str(list3) + "）不一致", "red")
        else:
            table_father.display(self, "检查（勘验）人（签名）：检查（勘验）人（签名）：和日期一致", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, '检查（勘验）人（签名）：', '主观审查是否为空')

    def xianchangqingkuang_gaozhishixiang_case1(self):
        pattern = ".*告知事项：([\s\S]*).*现场情况"
        gaozhishixiang = re.findall(pattern, self.contract_text)
        if gaozhishixiang == [] or gaozhishixiang[0].strip() == "":
            table_father.display(self, "告知事项：告知事项 不能为空", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, '告知事项：', '告知事项 不能为空')
        else:
            gaozhishixiang = gaozhishixiang[0].replace("\n", "").replace(" ", "")
            # print(gaozhishixiang)
            pattern = ".*依法对(.*)进行检查，请予以配合.*"
            x1 = re.search(pattern, gaozhishixiang)
            if x1 == None:
                table_father.display(self, "告知事项：告知事项 中 未说明来意", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '告知事项：', '告知事项 中 未说明来意')
            else:
                table_father.display(self, "告知事项：告知事项 中 说明来意", "green")

        pattern = ".*现场情况：([\s\S]*).*被检查（勘验）人或现场负责人（签名）"
        xiangchangqingkuang = re.findall(pattern, self.contract_text)
        if xiangchangqingkuang == [] or xiangchangqingkuang[0].strip() == "":
            table_father.display(self, "现场情况：现场情况 不能为空", "green")
            tyh.addRemarkInDoc(self.mw, self.doc, '现场情况：', '现场情况 不能为空')
        else:
            xiangchangqingkuang = xiangchangqingkuang[0].replace("\n", "").replace(" ", "")
            # print(xiangchangqingkuang)

            # 文本中的日期应与检查（勘验）时间的开始时间一致
            # print(tyh.jiancha_time(self.source_prifix)[0][0])
            jiancha_time = tyh.jiancha_time(self.source_prifix)
            if jiancha_time != False:
                if jiancha_time[0][0] not in xiangchangqingkuang:
                    table_father.display(self, "现场情况：“现场情况”中的日期（" + str(
                        xiangchangqingkuang) + "）应与“检查（勘验）时间”的开始时间（" + str(jiancha_time[0][0]) + "）不一致",
                                         "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '现场情况：', "“现场情况”中的日期（" + str(
                        xiangchangqingkuang) + "）应与“检查（勘验）时间”的开始时间（" + str(jiancha_time[0][0]) + "）不一致")
                else:
                    table_father.display(self, "现场情况：现场情况 中的日期应与检查（勘验）时间的开始时间 一致", "green")
            else:
                table_father.display(self, "现场情况：检查（勘验）时间 无法找到", "red")

            # 文本中的地点应与检查（勘验）地点一致
            jiancha_place = tyh.jiancha_place(self.source_prifix)
            if jiancha_place != False:
                if jiancha_place[0] not in xiangchangqingkuang:
                    table_father.display(self, "现场情况：“现场情况”中的地点（" + str(
                        xiangchangqingkuang) + "）应与“检查（勘验）地点”（" + str(jiancha_place[0]) + "）不一致", "red")
                    tyh.addRemarkInDoc(self.mw, self.doc, '现场情况：', "“现场情况”中的地点（" + str(
                        xiangchangqingkuang) + "）应与“检查（勘验）地点”（" + str(jiancha_place[0]) + "）不一致")
                else:
                    table_father.display(self, "现场情况：现场情况 中的地点应与检查（勘验）地点 一致", "green")
            else:
                table_father.display(self, "现场情况：检查（勘验）地点 无法找到", "red")

            pattern = ".*(共计(.*)品种(.*)卷烟).*"
            x1 = re.search(pattern, xiangchangqingkuang)
            if x1 == None:
                table_father.display(self, "现场情况：现场情况 中 未含有卷烟数目要素", "red")
                tyh.addRemarkInDoc(self.mw, self.doc, '现场情况：', '现场情况 中 未含有卷烟数目要素')
            else:
                table_father.display(self, "现场情况：现场情况 中 含有卷烟数目要素", "green")

        # 文本中是否包含出示证件、说明来意、卷烟数目、法律条款等要素
        pattern = ".*出示(.*)证件.*"
        x1 = re.search(pattern, gaozhishixiang)
        x2 = re.search(pattern, xiangchangqingkuang)
        if x1 == None and x2 == None:
            table_father.display(self, "告知事项：告知事项与现场情况 中 未含有出示证件要素", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '告知事项：', '告知事项与现场情况 中 未含有出示证件要素')
        else:
            table_father.display(self, "告知事项：告知事项与现场情况 中 含有出示证件要素", "green")

        # pattern = ".*(《.*》).*"
        # x1 = re.search(pattern, gaozhishixiang)
        # x2 = re.search(pattern, xiangchangqingkuang)
        # if x1 == None and x2 == None:
        if "烟草专卖零售许可证" not in gaozhishixiang and "烟草专卖零售许可证" not in xiangchangqingkuang:
            table_father.display(self, "告知事项：告知事项与现场情况 中 未含有法律条款要素", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '告知事项：', '告知事项与现场情况 中 未含有法律条款要素')
        else:
            table_father.display(self, "告知事项：告知事项与现场情况 中 含有法律条款要素", "green")

        pattern = ".*(行政执法人员.*?[（，]).*"
        x1 = re.findall(pattern, gaozhishixiang)
        people1 = self.entityrecognition.get_identity_with_tag(x1[0], 'PER')

        x2 = re.findall(pattern, xiangchangqingkuang)
        people2 = self.entityrecognition.get_identity_with_tag(x2[0], 'PER')

        if (len(people1) < 2):
            table_father.display(self, "告知事项：告知事项文本中的执法人员 不足两人", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '告知事项：', '告知事项文本中的执法人员 不足两人')
        else:
            table_father.display(self, "告知事项：告知事项文本中的执法人员>=两人", "green")

        if (len(people2) < 2):
            table_father.display(self, "告知事项：现场情况文本中的执法人员 不足两人", "red")
            tyh.addRemarkInDoc(self.mw, self.doc, '现场情况：', '现场情况文本中的执法人员 不足两人')
        else:
            table_father.display(self, "告知事项：现场情况文本中的执法人员>=两人", "green")


if __name__ == '__main__':
    my_prefix = "C:\\Users\\Zero\\OneDrive\\案卷\\tyh\\"
    list = os.listdir(my_prefix)
    if "检查（勘验）笔录_.docx" in list:
        ioc = table6(my_prefix, my_prefix)
        contract_file_path = my_prefix + "检查（勘验）笔录_.docx"
        ioc.check(contract_file_path, "检查（勘验）笔录_.docx")
