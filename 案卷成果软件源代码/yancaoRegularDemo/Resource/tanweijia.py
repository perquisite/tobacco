import os
#from tools.utils import get_root_dir
import shutil
from yancaoRegularDemo.Resource.table_37_54.table38 import Table_38
from yancaoRegularDemo.Resource.table_37_54.table39 import Table_39
from yancaoRegularDemo.Resource.table_37_54.table40 import Table_40
from yancaoRegularDemo.Resource.table_37_54.table41 import Table_41
from yancaoRegularDemo.Resource.table_37_54.table52 import Table_52
from yancaoRegularDemo.Resource.table_37_54.table53 import Table_53
from yancaoRegularDemo.Resource.table_37_54.table54 import Table_54


class Precessor3:
    # 传入单个目标文件路径，传入导出目标文件夹的位置
    # 此处，应该先根据【源文件路径】，拷贝一份文件到【目标文件夹】
    def __init__(self, file_path, export_dir=None, need_to_export=None):
        self.file_path = file_path
        if need_to_export or need_to_export is None:
            # 默认需要导出到 【目标文件夹】，拷贝一份源文件到目标文件夹
            shutil.copy(file_path, export_dir + '/' + file_path[file_path.rfind('/', 1) + 1:])
            self.my_prefix = export_dir + '/'
        else:
            # 不需要导出到 【目标文件夹】，直接对源文件夹进行处理
            self.my_prefix = file_path[0:file_path.rfind('/', 1) + 1]
            # self.export_dir = export_dir

    def action(self):
        # list_all = os.listdir(my_prefix)
        # if "撤销立案报告表_.docx" in list_all:
        #     ioc = Table19(my_prefix)
        #     contract_file_path = my_prefix + "撤销立案报告表_.docx"
        #     ioc.check(contract_file_path)
        # if "案件调查终结报告_.docx" in list_all:
        #     ioc = Table20(my_prefix)
        #     contract_file_path = my_prefix + "案件调查终结报告_.docx"
        #     ioc.check(contract_file_path)
        # if "延长调查终结审批表_.docx" in list_all:
        #     ioc = Table21(my_prefix)
        #     contract_file_path = my_prefix + "延长调查终结审批表_.docx"
        #     ioc.check(contract_file_path)
        # if "延长调查期限告知书_.docx" in list_all:
        #     ioc = Table22(my_prefix)
        #     contract_file_path = my_prefix + "延长调查期限告知书_.docx"
        #     ioc.check(contract_file_path)
        # if "先行登记保存证据处理通知书_.docx" in list_all:
        #     ioc = Table23(my_prefix)
        #     contract_file_path = my_prefix + "先行登记保存证据处理通知书_.docx"
        #     ioc.check(contract_file_path)
        # if "涉案物品返还清单_.docx" in list_all:
        #     ioc = Table24(my_prefix)
        #     contract_file_path = my_prefix + "涉案物品返还清单_.docx"
        #     ioc.check(contract_file_path)
        # if "行政处罚事先告知书_.docx" in list_all:
        #     ioc = Table25(my_prefix)
        #     contract_file_path = my_prefix + "行政处罚事先告知书_.docx"
        #     ioc.check(contract_file_path)
        # if "陈述申辩记录_.docx" in list_all:
        #     ioc = Table26(my_prefix)
        #     contract_file_path = my_prefix + "陈述申辩记录_.docx"
        #     ioc.check(contract_file_path)
        # if "听证告知书_.docx" in list_all:
        #     ioc = Table27(my_prefix)
        #     contract_file_path = my_prefix + "听证告知书_.docx"
        #     ioc.check(contract_file_path)
        # if "听证通知书_.docx" in list_all:
        #     ioc = Table28(my_prefix)
        #     contract_file_path = my_prefix + "听证通知书_.docx"
        #     ioc.check(contract_file_path)
        # if "不予受理听证通知书_.docx" in list_all:
        #     ioc = Table29(my_prefix)
        #     contract_file_path = my_prefix + "不予受理听证通知书_.docx"
        #     ioc.check(contract_file_path)
        # if "听证公告_.docx" in list_all:
        #     ioc = Table30(my_prefix)
        #     contract_file_path = my_prefix + "听证公告_.docx"
        #     ioc.check(contract_file_path)
        # if "听证笔录_.docx" in list_all:
        #     ioc = Table31(my_prefix)
        #     contract_file_path = my_prefix + "听证笔录_.docx"
        #     ioc.check(contract_file_path)
        # if "听证报告_.docx" in list_all:
        #     ioc = Table32(my_prefix)
        #     contract_file_path = my_prefix + "听证报告_.docx"
        #     ioc.check(contract_file_path)
        # if "案件集体讨论记录_.docx" in list_all:
        #     ioc = Table33(my_prefix)
        #     contract_file_path = my_prefix + "案件集体讨论记录_.docx"
        #     ioc.check(contract_file_path)
        # --------------以上为旧版本，可用于系统测试--------------

        # ioc.check 返回提示信息，为此，需为每个表增加一个存储提示信息的list
        # 此处需要注意，表增加了一个file_path参数，因为一个表中可能需要打开其他的表，此时需要用到源路径，而不是目标路径
        if "不予行政处罚决定书_.docx" in self.file_path:
            ioc = Table_38(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result

        if "送达回证_.docx" in self.file_path:
            ioc = Table_39(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result

        if "送达公告_.docx" in self.file_path:
            ioc = Table_40(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result

        if "责令改正通知书_.docx" in self.file_path:
            ioc = Table_41(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result

        if "卷宗封面_.docx" in self.file_path:
            ioc = Table_52(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result

        if "卷宗目录_.docx" in self.file_path:
            ioc = Table_53(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result

        if "卷内备考表_.docx" in self.file_path:
            ioc = Table_54(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result

        else:
            return ["不存在表"]

# if __name__ == '__main__':
    # my_prefix = "D:\烟草项目\\tobacco-xiejunyu\\tobacco-xiejunyu\yancaoRegularDemo\副本\\"
    # # print(my_prefix)
    # ioc = Table_39(my_prefix)
    # contract_file_path = my_prefix + "送达回证_.docx"
    # # print(contract_file_path)
    # ioc.check(contract_file_path)
    # my_prefix = r"D:/烟草项目\tobacco-8.7/yancaoRegularDemo/副本/责令改正通知书_.docx"
    # print("\n")
    # print(my_prefix)
    # ioc = Precessor3(my_prefix, None, False)
    # ioc.action()