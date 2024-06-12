import sys
import os
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)

import re
import time
from zhishitupu.src.function import *
from zhishitupu.tools.check_excel import check_excel
from zhishitupu.tools.transform import excle_to_rdf

template_src_path = getRootPath() + "模板.xls"

# 1、“查询”相关
"""
    （函数1）
    无参数输入，直接返回一个字典，包含了各个市的条目名称，例如{‘成都市局’：[‘1’,‘2’,‘3’], ‘德阳’：[‘1’,‘2’],…}
"""


def show_city_and_clause():
    rdf_folder_path = getRootPath() + "\\" + 'rdf'
    file_paths = []
    output_dict = {}
    for root, dirs, files in os.walk(rdf_folder_path):
        for d in dirs:
            p = rdf_folder_path + "\\" + d
            for root, dirs, files in os.walk(p):
                for f in files:
                    file_name = f.rstrip(".rdf")
                    file_paths.append(file_name)
            if file_paths:
                output_dict[d] = file_paths
            file_paths = []
    return output_dict

# 此函数内部调用，不是api函数
def show_city_and_clause_1():
    rdf_folder_path = getRootPath() + "\\" + 'rdf'
    file_paths = []
    output_dict = {}
    for root, dirs, files in os.walk(rdf_folder_path):
        for d in dirs:
            p = rdf_folder_path + "\\" + d
            for root, dirs, files in os.walk(p):
                for f in files:
                    file_name = f.rstrip(".rdf")
                    file_paths.append(file_name)
            output_dict[d] = file_paths
            file_paths = []
    return output_dict


"""
    （函数2）
    输入是某市某条,例如（‘成都市局’，'1'）），类型都是字符串，若存在该条，则输出是唤起的Excel表格,返回TRUE；
    若不存在，则输出是报错提示字符串。
"""


def query(city, clause):
    info_dict = show_city_and_clause()
    if city in info_dict.keys():  # 含有该城市
        clause_list = info_dict[city]
        if str(clause) in clause_list:  # 且含有该条
            cache_path = getRootPath() + "缓存文件夹" + "\\" + city + "第" + str(clause) + "条.xls"
            os.startfile(cache_path)
            time.sleep(0.2)
            return True
        else:
            return '该城市对应库中无该法条。您可以先尝试添加该条。'
    else:
        return '库中无此城市的相关信息。您可以先尝试添加信息。'


# 2、“增加”相关
"""
    （函数3）
    输入是某市某条,例如（‘成都市局’，'1'）），类型都是字符串。若存在该条，则输出是唤起空的的Excel表格（该Excel文档在项目内一个固定的缓存文件夹）
    并返回该路径；若不存在，则输出是报错提示字符串。
"""


def add(city, clause):
    info_dict = show_city_and_clause_1()
    template_dst_path = getRootPath() + r"缓存文件夹" + "\\" + city + "第" + str(clause) + "条" + ".xls"
    if city in info_dict.keys():  # 含有该城市
        clause_list = info_dict[city]
        if str(clause) in clause_list:  # 且含有该条
            return '该法条信息已存在。您可以尝试修改该条。'
        else:
            # 创建（不用创建对应城市rdf文件夹）——拷贝空白模板到‘缓存文件夹’，并打开，返回拷贝到‘缓存文件夹’的Excel文件路径
            # template_dst_path = getRootPath() + r"缓存文件夹" + "\\" + city + "第" + str(clause) + "条" + ".xls"  # 缓存文件夹的路径名
            try:
                shutil.copyfile(template_src_path, template_dst_path)  # 先copy到缓存文件夹里
                os.startfile(template_dst_path)  # 换气
            except Exception as e:
                print('Excel表添加函数add:' + '遇到了一个错误:' + str(e.args))
            time.sleep(0.2)
            return template_dst_path
    else:
        # 创建（要先创建对应城市rdf文件夹）——先创建对应城市rdf文件夹，拷贝空白模板到‘缓存文件夹’，
        # 并打开，返回拷贝到‘缓存文件夹’的Excel文件路径
        city_path = getRootPath() + "rdf" + "\\" + city  # rdf库中的城市文件夹路径
        try:
            os.mkdir(city_path)  # 创建对应城市文件夹
            # template_dst_path = getRootPath() + r"缓存文件夹" + "\\" + city + "第" + str(clause) + "条" + ".xls"
            shutil.copyfile(template_src_path, template_dst_path)  # 先copy到缓存文件夹里
            os.startfile(template_dst_path)
        except Exception as e:
            print('Excel表添加函数 add:' + ' has occurred an error:' + str(e.args))
        time.sleep(0.2)
        return template_dst_path


"""
    （函数4）
    输入是要转化的Excel文档路径，它会执行对Excel表的检查工作，若表格无错误，则后端进行转化等工作并返回True；
     若表格有错误，则返回报错信息的字符串的list。
"""


def check_and_transform(excel_path):
    result_list = check_excel(excel_path)
    if result_list:  # 不是空list——含有报错信息，不能转化
        return result_list
    else:  # 空list——不含有报错信息，转化
        temp = re.search(r"缓存文件夹\\(\S+)第", excel_path)
        city_name = temp.group(1)
        temp = re.search(r"第(\S+)条", excel_path)
        clause = temp.group(1)
        save_path = getRootPath() + "rdf" + "\\" + city_name + "\\" + clause + ".rdf"  # 转化后的路径
        #print("excel_path:  " + excel_path)
        #print("save_path:  " + save_path)
        try:
            excle_to_rdf(excel_path, save_path)
            return True
        except Exception as e:
            print('Excel表转化函数 excel_to_rdf:' + ' has occurred an error:' + str(e.args))


# 3、“删除”相关
"""
    （函数5）
    输入某市某条（例如（‘成都市局’，‘1’））并点击“删除”后，若存在该条，则删除并返回True；若不存在，则输出是报错提示字符串。
"""


def delete(city, clause):
    should_delete_rdf_path = getRootPath() + "rdf" + "\\" + city + "\\" + str(clause) + ".rdf"
    cache_path = getRootPath() + "缓存文件夹" + "\\" + city + "第" + str(clause) + "条.xls"
    #print("delete_rdf_path + " + should_delete_rdf_path)
    #print("cache_path + " + should_delete_rdf_path)
    info_dict = show_city_and_clause()
    if city in info_dict.keys():  # 含有该城市
        clause_list = info_dict[city]
        if str(clause) in clause_list:  # 且含有该条 则删除
            try:
                os.remove(should_delete_rdf_path)  # 删除rdf文件
                os.remove(cache_path)  # 删除缓存的Excel文件
                return True
            except Exception as e:
                print('删除失败' + ' 错误信息:' + str(e.args))
        else:
            return "该城市的库中不存在此法条。您可以尝试先添加。"
    else:
        return "库中不含该城市的相关信息。您可以尝试先添加。"


# 4、“修改”相关
"""
    （函数6）
    输入某市某条如（‘成都市局’，‘1’））并点击“更改”后，返回报错提示字符串（结束，因为没有该条存档的情况下无法更改），
    否则是唤起缓存的含有之前内容Excel文档，返回True。（用户填完后，还要调用函数4再生成一遍）
"""


def edit(city, clause):
    info_dict = show_city_and_clause()
    if city in info_dict.keys():  # 含有该城市
        clause_list = info_dict[city]
        if str(clause) in clause_list:  # 且含有该条 则进行修改
            cache_path = getRootPath() + "缓存文件夹" + "\\" + city + "第" + str(clause) + "条.xls"
            try:
                os.startfile(cache_path)  # 唤起excel文件
                time.sleep(0.2)
                return cache_path
            except Exception as e:
                print('修改失败' + ' 错误信息:' + str(e.args))
        else:
            return "该城市的库中不存在此法条。您可以尝试先添加。"
    else:
        return "库中不含该城市的相关信息。您可以尝试先添加。"

"""
格式： python kg_api.py [操作] (([城市] [条数])或(路径))
[操作]：1、get: 就一条命令：python kg_api.py get
       2、query: 例:python kg_api.py query 成都市局 1
       3、add: 例:python kg_api.py add 成都市局 1
       4、delete: 例:python kg_api.py delete 成都市局 1
       5、edit: 例:python kg_api.py edit 成都市局 1
       6、transform：格式:python kg_api.py edit transform [要检查并转化的Excel路径]
"""
if __name__ == '__main__':
    args = sys.argv
    operation = args[1]
    if operation == "get":  # 函数1
        print(show_city_and_clause())
    elif operation == "transform":  # 函数4
        excelpath = args[2]
        print(check_and_transform(excelpath))
    else:
        city = args[2]
        clause = args[-1]
        if operation == "query":  # 函数2
            print(query(city, clause))
        elif operation == "add":  # 函数3
            print(add(city, clause))
        elif operation == "delete":  # 函数5
            print(delete(city, clause))
        elif operation == "edit":  # 函数6
            print(edit(city, clause))
