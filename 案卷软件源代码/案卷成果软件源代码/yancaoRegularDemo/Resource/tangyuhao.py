import os
import shutil

import docx

from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father
from yancaoRegularDemo.Resource.table_1_18.table1 import table1
from yancaoRegularDemo.Resource.table_1_18.table2 import table2
from yancaoRegularDemo.Resource.table_1_18.table3 import table3
from yancaoRegularDemo.Resource.table_1_18.table4 import table4
from yancaoRegularDemo.Resource.table_1_18.table5 import table5
from yancaoRegularDemo.Resource.table_1_18.table6 import table6
from yancaoRegularDemo.Resource.table_1_18.table7 import table7
from yancaoRegularDemo.Resource.table_1_18.table8 import table8
from yancaoRegularDemo.Resource.table_1_18.table9 import table9
from yancaoRegularDemo.Resource.table_1_18.table10 import table10
from yancaoRegularDemo.Resource.table_1_18.table11 import table11
from yancaoRegularDemo.Resource.table_1_18.table12 import table12
from yancaoRegularDemo.Resource.table_1_18.table13 import table13
from yancaoRegularDemo.Resource.table_1_18.table14 import table14
from yancaoRegularDemo.Resource.table_1_18.table15 import table15
from yancaoRegularDemo.Resource.table_1_18.table16 import table16
from yancaoRegularDemo.Resource.table_1_18.table17 import table17
from yancaoRegularDemo.Resource.table_1_18.table18 import table18

from yancaoRegularDemo.Resource.table_19_36.Table19 import Table19
from yancaoRegularDemo.Resource.table_19_36.Table20 import Table20
from yancaoRegularDemo.Resource.table_19_36.Table21 import Table21
from yancaoRegularDemo.Resource.table_19_36.Table22 import Table22
from yancaoRegularDemo.Resource.table_19_36.Table23 import Table23
from yancaoRegularDemo.Resource.table_19_36.Table24 import Table24
from yancaoRegularDemo.Resource.table_19_36.Table25 import Table25
from yancaoRegularDemo.Resource.table_19_36.Table26 import Table26
from yancaoRegularDemo.Resource.table_19_36.Table27 import Table27
from yancaoRegularDemo.Resource.table_19_36.Table28 import Table28
from yancaoRegularDemo.Resource.table_19_36.Table29 import Table29
from yancaoRegularDemo.Resource.table_19_36.Table30 import Table30
from yancaoRegularDemo.Resource.table_19_36.Table31 import Table31
from yancaoRegularDemo.Resource.table_19_36.Table32 import Table32
from yancaoRegularDemo.Resource.table_19_36.Table33 import Table33
from yancaoRegularDemo.Resource.table_19_36.Table34 import Table34
from yancaoRegularDemo.Resource.table_19_36.Table35 import Table35
from yancaoRegularDemo.Resource.table_19_36.Table36 import Table36
from yancaoRegularDemo.Resource.table_1_18.条形码 import table0
from yancaoRegularDemo.Resource.table_37_54.table37 import Table37

from yancaoRegularDemo.Resource.table_37_54.table38 import Table_38
from yancaoRegularDemo.Resource.table_37_54.table39 import Table_39
from yancaoRegularDemo.Resource.table_37_54.table40 import Table_40
from yancaoRegularDemo.Resource.table_37_54.table41 import Table_41
from yancaoRegularDemo.Resource.table_37_54.table44 import Table44
from yancaoRegularDemo.Resource.table_37_54.table51 import Table51
from yancaoRegularDemo.Resource.table_37_54.table52 import Table_52
from yancaoRegularDemo.Resource.table_37_54.table53 import Table_53
from yancaoRegularDemo.Resource.table_37_54.table54 import Table_54

from yancaoRegularDemo.Resource.table_37_54.table60 import Table_60
from yancaoRegularDemo.Resource.tools.get_pictures import get_pictures_single


class Precessor1:
    # 传入单个目标文件路径，传入导出目标文件夹的位置
    # 此处，应该先根据【源文件路径】，拷贝一份文件到【目标文件夹】
    def __init__(self, file_path, export_dir=None, need_to_export=None):
        file_path=file_path.replace("\\", "/")
        self.file_path = file_path
        self.total_count = 0
        self.wrong_count = 0
        export_dir = export_dir.replace("\\", "/")



        if need_to_export or need_to_export is None:
            # 默认需要导出到 【目标文件夹】，拷贝一份源文件到目标文件夹
            object_path=export_dir + '/' + file_path[file_path.rfind('/', 1) + 1:]
            shutil.copy(file_path, object_path)
            self.my_prefix = export_dir + '/'
        else:
            # 不需要导出到 【目标文件夹】，直接对源文件夹进行处理
            self.my_prefix = file_path[0:file_path.rfind('/', 1) + 1]
            # self.export_dir = export_dir
        if ":/" in self.my_prefix:
            self.my_prefix = self.my_prefix.replace(":/", ":\\")

    def action(self):
        # ioc.check 返回提示信息，为此，需为每个表增加一个存储提示信息的list
        # 此处需要注意，表增加了一个file_path参数，因为一个表中可能需要打开其他的表，此时需要用到源路径，而不是目标路径
        #
        #try:
            if "举报记录表" in self.file_path:
                ioc = table1(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "立案报告表" in self.file_path and "不予立案报告表" not in self.file_path and "撤销立案报告表" not in self.file_path:
                ioc = table2(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "不予立案报告表" in self.file_path:
                ioc = table2(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "延长立案期限审批表" in self.file_path:
                ioc = table3(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "延长立案期限告知书" in self.file_path:
                ioc = table4(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "指定管辖通知书" in self.file_path:
                ioc = table5(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "检查（勘验）笔录" in self.file_path:
                ioc = table6(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "先行登记保存批准书" in self.file_path:
                ioc = table7(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "证据先行登记保存通知书" in self.file_path:
                ioc = table8(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "抽样取证物品清单" in self.file_path:
                ioc = table9(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "询问笔录" in self.file_path:
                ioc = table10(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "涉案烟草专卖品核价表" in self.file_path:
                ioc = table11(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "证据复制提取单" in self.file_path:
                ioc = table12(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "公告" in self.file_path and "送达公告" not in self.file_path:
                ioc = table13(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "卷烟鉴别检验样品留样、损耗费用审批表" in self.file_path:
                ioc = table14(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "案件移送函" in self.file_path:
                ioc = table15(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "案件移送回执" in self.file_path:
                ioc = table16(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "移送财物清单" in self.file_path:
                ioc = table17(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "协助调查函" in self.file_path:
                ioc = table18(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "条形码" in self.file_path:
                ioc = table0(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            # xiejunyu
            if "撤销立案报告表" in self.file_path:
                ioc = Table19(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "调查终结报告" in self.file_path:
                ioc = Table20(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "延长调查终结" in self.file_path:
                ioc = Table21(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "延长调查期限告知书" in self.file_path:
                ioc = Table22(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "先行登记保存证据处理通知书" in self.file_path:
                ioc = Table23(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "涉案物品返还清单" in self.file_path:
                ioc = Table24(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "行政处罚事先告知" in self.file_path:
                ioc = Table25(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "陈述申辩记录" in self.file_path:
                ioc = Table26(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "听证告知书" in self.file_path:
                ioc = Table27(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "听证通知书" in self.file_path:
                ioc = Table28(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "不予受理听证通知书" in self.file_path:
                ioc = Table29(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "听证公告" in self.file_path:
                ioc = Table30(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "听证笔录" in self.file_path:
                ioc = Table31(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "听证报告" in self.file_path:
                ioc = Table32(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "案件集体讨论记录" in self.file_path:
                ioc = Table33(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "案件处理审批表" in self.file_path:
                ioc = Table34(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "当场行政处罚决定书" in self.file_path:
                ioc = Table35(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            if "行政处罚决定书" in self.file_path and "不予行政处罚决定书" not in self.file_path:
                ioc = Table36(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            # xiejunyu
            # #
            # tanweijia

            if "行政处理决定书" in self.file_path:
                ioc = Table37(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            if "不予行政处罚决定书" in self.file_path:
                ioc = Table_38(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            if "送达回证" in self.file_path:
                ioc = Table_39(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            if "送达公告" in self.file_path:
                ioc = Table_40(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            if "责令改正通知书" in self.file_path:
                ioc = Table_41(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            if "违法物品销毁记录表" in self.file_path:
                ioc = Table44(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            if "结案报告表" in self.file_path:
                ioc = Table51(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            if "卷宗封面" in self.file_path:
                ioc = Table_52(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            if "卷宗目录" in self.file_path:
                ioc = Table_53(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            if "卷内备考表" in self.file_path:
                ioc = Table_54(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result

            if "含身份证文书" in self.file_path:
                ioc = Table_60(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
                info_list_result = ioc.check(self.file_path, self.file_path[self.file_path.rfind('/', 1) + 1:])
                self.total_count, self.wrong_count = ioc.get_count()
                return info_list_result
            # tanweijia
            else:
                return ["不存在表\n-----------------------------------"]
        # except Exception as e:
        #     table_father.display(self, 'Precessor1 has occurred an error:' + str(e.args))

    def processor1_get_count(self):
        return self.total_count, self.wrong_count
