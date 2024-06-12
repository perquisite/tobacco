import shutil
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


# my_prefix = get_root_dir() + "\\yancaoRegularDemo\\副本\\"

class Precessor2:
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
        # ioc.check 返回提示信息，为此，需为每个表增加一个存储提示信息的list
        # 此处需要注意，表增加了一个file_path参数，因为一个表中可能需要打开其他的表，此时需要用到源路径，而不是目标路径
        if "撤销立案报告表_.docx" in self.file_path:
            ioc = Table19(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "案件调查终结报告_.docx" in self.file_path:
            ioc = Table20(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "延长调查终结审批表_.docx" in self.file_path:
            ioc = Table21(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "延长调查期限告知书_.docx" in self.file_path:
            ioc = Table22(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "先行登记保存证据处理通知书_.docx" in self.file_path:
            ioc = Table23(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "涉案物品返还清单_.docx" in self.file_path:
            ioc = Table24(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "行政处罚事先告知书_.docx" in self.file_path:
            ioc = Table25(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "陈述申辩记录_.docx" in self.file_path:
            ioc = Table26(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "听证告知书_.docx" in self.file_path:
            ioc = Table27(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "听证通知书_.docx" in self.file_path:
            ioc = Table28(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "不予受理听证通知书_.docx" in self.file_path:
            ioc = Table29(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "听证公告_.docx" in self.file_path:
            ioc = Table30(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "听证笔录_.docx" in self.file_path:
            ioc = Table31(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "听证报告_.docx" in self.file_path:
            ioc = Table32(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "案件集体讨论记录_.docx" in self.file_path:
            ioc = Table33(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "案件处理审批表_.docx" in self.file_path:
            ioc = Table34(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "当场行政处罚决定书_.docx" in self.file_path:
            ioc = Table35(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        if "行政处罚决定书_.docx" in self.file_path:
            ioc = Table36(self.my_prefix, self.file_path[0:self.file_path.rfind('/', 1) + 1])
            info_list_result = ioc.check(self.file_path)
            return info_list_result
        else:
            return ["不存在表"]
