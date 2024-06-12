import os
import shutil
#from tools.transform import excle_to_rdf

# def add_and_cleanCache(template_dst_path, create_rdf_path):  # 转化添加rdf并清除缓存文件
#     excle_to_rdf(template_dst_path, create_rdf_path)
#     os.remove(template_dst_path)  # 清除缓存文件 （万一不关闭，需要报错处理）


def getRootPath():
    # 获取文件目录
    curPath = os.path.abspath(os.path.dirname(__file__))
    # 获取项目根路径，内容为当前项目的名字
    rootPath = curPath[:curPath.find("zhishitupu\\") + len("zhishitupu\\")]
    return rootPath


def get_allfile(path):  # 获取所有文件和文件夹
    all_file = []
    for f in os.listdir(path):  # listdir返回文件中所有目录
        f_name = os.path.join(path, f)
        all_file.append(f_name)
    return all_file


def copy_and_start(template_src_path, template_dst_path):  # 考虑报错？？
    shutil.copyfile(template_src_path, template_dst_path)  # 先copy到缓存文件夹里
    os.startfile(template_dst_path)


def is_list_consecutive(lista):
    sorted_list = sorted(lista)
    #sorted(l) ==
    #range_list = list(range(min(lista), max(lista)+1))
    #print(sorted_list)
    #print(lista)
    if not sorted_list == lista:
        return False
    else:
        i = 0
        while i < len(lista)-1:
            if (lista[i] + 1) == lista[i+1]:
                i += 1
                continue
            else:
                return False
        return True


if __name__ == '__main__':
    print(getRootPath())
    # template_src_path = r"C:\Users\twj\PycharmProjects\KnowledgeGraph\模板.xlsx"
    # template_dst_path = r"C:\Users\twj\PycharmProjects\KnowledgeGraph\缓存文件夹"
    # shutil.copyfile(template_src_path, template_dst_path)
