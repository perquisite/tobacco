"""
    Project:中国烟草案卷执法组文件操作类
    Author:陈付旻
    Date:2021-06-26 18:30
"""


class FileOperator(object):
    def __init__(self):
        super(FileOperator, self).__init__()

    """
        两个文书名称，一个是你获取到的，一个是你要判断的文书
    """
    def fileIsExist(self, targetName, sourceName):
        if targetName == sourceName:
            print(sourceName+'文书存在！')
        else:
            print(sourceName+'文书不存在！')

