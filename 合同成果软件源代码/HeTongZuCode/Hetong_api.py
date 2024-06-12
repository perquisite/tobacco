import json
import sys
import requests as req


# file_paths: 要处理文件的路径list export_path：保存文件夹路径
# def deal_files(files, export_path):
#     data = {
#         'files': files,
#         'export_path': export_path
#     }
#     r = req.post(url='http://127.0.0.1:9001/deal_file', data=json.dumps(data))
#     print(r.text)


# files: 要处理文件的路径,标准,类型list export_path：保存文件夹路径
from ContractType import ContractType
from DocReader import DocReader
from over_all_description import get_over_all_file

import pythoncom
from win32com.client import Dispatch
def deal_files(files, export_path):
    # try:
    #     pythoncom.CoInitialize()
    #     word = Dispatch('Word.Application')
    #     pythoncom.CoInitialize()
    #     word.Documents.close()
    # except:
    #     print("关闭word进程")

    over_all_info = []
    export_path = export_path.replace('\\', '/')
    print(files)
    try:
        for file_path in files:

            try:
                file_path[0] = file_path[0].replace('\\', '/')
                file_path[2] = int(file_path[2])
                print(file_path)
                d = DocReader(file_path[0], export_path)
                contract_check_result = d.deal_one(file_path[2], file_path[1])
                if len(contract_check_result.factors.keys()) == 0:
                    print("合同审查出错！")
                if contract_check_result.type == ContractType.NotSure:
                    print("合同为非标准合同，对该合同进行风险性检测和完整性检测！")
                if contract_check_result != "docx_blank" and contract_check_result != "type_not_sure" and contract_check_result.type != ContractType.NotSure:
                    print("错误的要素:\n" + str(contract_check_result.factors_error))
                    print("提示的要素:\n" + str(contract_check_result.factors_to_inform))
            except Exception as e:
                print('合同出错：'.format(file_path[0]))
            finally:
                over_all_info.append([file_path[0], contract_check_result])
                print('审查成功！')
        get_over_all_file(over_all_info, export_path)
        return True
    except Exception as e:
        print('合同审查出现错误{}'.format(e))
        get_over_all_file([], export_path)
        return False


# 命令： python Hetong_api.py file1  a b file2 a b file3 a b export_path
# 一定要绝对路径，且文件名不能有空格
if __name__ == '__main__':
    args = sys.argv
    # args.extend([r'C:\Users\Zero\Desktop\合同组测试用例\合同组测试用例\非标准文件\采购合同\非标准采购合同(1).docx', 'nonstandard', '2', r'C:\Users\Zero\Desktop'])
    arg = args[1:-1]
    files = []
    for i in range(len(arg) // 3):
        files.append(arg[i * 3: (i + 1) * 3])
        files[i][0] = files[i][0].replace('\'','')
    export_path = args[-1]
    print("处理：{}个文件".format(len(files)))
    deal_files(files, export_path)
