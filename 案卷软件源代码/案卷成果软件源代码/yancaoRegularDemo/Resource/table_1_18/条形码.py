from pyzbar.pyzbar import decode
from PIL import Image
from yancaoRegularDemo.Resource.tools.tangyuhao_readfile import *

import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.tools.utils import *
from yancaoRegularDemo.Resource.tools.simple_content import Simple_Content

from yancaoRegularDemo.Resource.tools.utils import is_valid_date

import win32com.client

from yancaoRegularDemo.Resource.ReadFile import *
import re
from yancaoRegularDemo.Resource.table_1_18.Table_Father import table_father

from warnings import simplefilter
from yancaoRegularDemo.Resource.tools.get_pictures import get_pictures_multi


class table0(table_father):
    def __init__(self, my_prefix, source_prifix):
        # super(table1, self).__init__()
        table_father.__init__(self)
        self.source_prifix = source_prifix
        self.contract_text = None
        self.contract_tables_content = None
        self.my_prefix = my_prefix

        self.all_to_check = [
            "self.pictureRight(file_name_real)"
        ]

    def pictureRight(self, file_name_real):
        global barcodeData
        picture_index = self.source_prifix + r"picture\\" + file_name_real.split(".")[0] + r"\word\media\image1.jpeg"

        barcodes = decode(Image.open(picture_index))

        for barcode in barcodes:
            barcodeData = barcode.data.decode("utf-8")
        table_father.display(self, barcodeData, "green")
        tyh.addRemarkInDoc(self.mw, self.doc, '条形码', barcodeData)

    def check(self, contract_file_path, file_name_real):
        print("正在审查" + file_name_real + "，审查结果如下：")
        self.mw = win32com.client.Dispatch("Word.Application")
        self.doc = self.mw.Documents.Open(self.my_prefix + file_name_real)
        data = DocxData(file_path=contract_file_path)
        self.contract_text = data.text
        self.contract_tables_content = data.tabels_content
        for func in self.all_to_check:
            try:
                eval(func)
            except Exception as e:
                table_father.display(self, "文档存在格式错误，函数失效：" + func + ' 遇到错误:' + str(e.args))
        self.doc.Close()
        # self.mw.Quit()
        print(file_name_real + "审查完毕\n")
        info_list_result = table_father.get_info_list(self)
        return info_list_result


if __name__ == '__main__':

    my_prefix = r"C:\Users\Zero\Desktop\副本\\"
    list = os.listdir(my_prefix)
    my_prefix = r"C:\Users\Zero\Desktop\副本\\"
    if "条形码_.docx" in list:
        ioc = table0(my_prefix, my_prefix)
        contract_file_path = os.path.join(my_prefix, "条形码_.docx")
        ioc.check(contract_file_path, "条形码_.docx")
