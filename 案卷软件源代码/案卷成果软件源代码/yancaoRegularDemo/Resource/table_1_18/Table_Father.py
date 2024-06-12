import os


class table_father(object):
    def __init__(self):
        self.info_list = []
        self.target_file_name_list = [] # 被比较文件名列表,与本类内is_target_file_exit()函数配合使用
        self.wrong_count = 0  # 处理的错误数目
        self.total_count = 0  # 处理的总数目
        self.manual_item_count = 0  # 需人工审查的数目

    def display(self, text, color=None):
        if color == 'red':
            self.info_list.append(text)
            self.wrong_count += 1
            self.total_count += 1
            # 以下用于console测试
            text = "\033[0;31m" + "warning: " + text + "\033[0m"
            print(text)
        elif color == "green":
            self.total_count += 1
            # self.info_list.append(text)，以下用于console测试
            text = "\033[0;36m" + text + "\033[0m"
            print(text)
        else:
            self.wrong_count += 1
            self.total_count += 1
            self.info_list.append(text)
            print(text)

    # 不适用于已经使用display函数进行add的类，主要用于MultiTableProcessor统计总数
    def get_count(self):
        return self.total_count, self.wrong_count

    def add_total_count(self, total_count):
        self.total_count += total_count

    def add_wrong_count(self, wrong_count):
        self.wrong_count += wrong_count

    def add_manual_item_count(self, manual_item_count):
        self.manual_item_count += manual_item_count

    def get_total_count(self):
        return self.total_count

    def get_wrong_count(self):
        return self.wrong_count

    def get_manual_item_count(self):
        return self.manual_item_count

    def get_info_list(self):
        return self.info_list

    def reset_count(self):
        self.total_count = 0
        self.wrong_count = 0
        self.manual_item_count = 0

    def is_target_file_exit(self, source_prifix, target_file_name_list):
        '''
        date：2022.2.21
        fuction:由于烟草命名存在不统一情况，此函数用于各个表子类中，当需要横向与其它文件进行比对时，判断被比较的docx文件是否存在。
        !!attention：子类需自行定义【被比较文件】拥有的命名list
        input:文件的目录前缀source_prifix；【被比较文件】可能拥有的命名list：target_file_name_list
        output:若目标文件不存在，返回-1。否则，返回list中对应的正确命名的下标
        '''
        if len(target_file_name_list) == 0:
            return -1
        for i, target_file_name in enumerate(target_file_name_list):
            if os.path.exists(source_prifix + target_file_name + ".docx") != 0:
                return i
            else:
                return -1
