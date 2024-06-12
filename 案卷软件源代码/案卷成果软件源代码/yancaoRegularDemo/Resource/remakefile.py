import os
import shutil
import win32com
from win32com.client import Dispatch


def remake(processed_file_sava_dir, filePath):
    # processed_file_sava_dir = r'D:\烟草\tobacco\yancaoRegularDemo\data\副本'
    # filePath = r'D:\烟草\tobacco\yancaoRegularDemo\data\案件处理审批表_ (6)'
    path = filePath
    filePath = processed_file_sava_dir + "\\" + filePath.split("\\")[-1] + ".docx"
    mw = win32com.client.Dispatch("Word.Application")
    doc = mw.Documents.Open(path)
    # 将word的数据保存到另一个文件
    doc.SaveAs(filePath, 12)
    # addRemarkInDoc(mw, doc, "身份证", "atttaaa")
    doc.SaveAs(filePath)
    doc.Close()


if __name__ == '__main__':
    source_file_dir = "C:\\Users\\Xie\\Desktop\\12.08案卷文书-程序分套"
    target_file_dir = "C:\\Users\\Xie\\Desktop\\out"
    for root, dirs, files in os.walk(source_file_dir):
        # for dir in dirs:
        #     print(os.path.join(root, dir))
        for file in files:
            print('------------------------------:root is:' + root)
            print(os.path.join(root, file))
            try:
                # 如果需要更改错误的文件格式,使用此功能,仅需修改source_file_dir,将自动遍历其中各级文件夹中的所有文档
                remake(root, os.path.join(root, file))

                # 如果需要删除特定文件,使用此功能
                # if '.docx' not in os.path.join(root, file):
                #     os.remove(os.path.join(root, file))

                # 如果需要移动若干同名文件，使用此功能,也会遍历其中各级文件夹中的所有文档
                # if '立案报告表' in os.path.join(root, file):
                #     shutil.copyfile(os.path.join(root, file), target_file_dir+'\\'+file)
            except Exception as e:
                print(e)
