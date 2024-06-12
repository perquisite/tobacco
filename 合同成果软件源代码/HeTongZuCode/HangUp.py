# -*- coding: utf-8 -*-
import json
import os
import sys

import pythoncom
from flask import Flask, request
from flask_cors import *
from DocReader import DocReader
from ContractType import ContractType
from over_all_description import get_over_all_file

app = Flask(__name__)


@app.route("/deal_file", methods=["POST"])
def deal_file():
    try:
        files = json.loads(request.data)['files']
        export_path = json.loads(request.data)['export_path']
        rs = deal_files(files, export_path)
    except Exception as e:
        print(e)
        rs = False
    return json.dumps({'status': rs})


# files: 要处理文件的路径,标准,类型list export_path：保存文件夹路径
def deal_files(files, export_path):
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
        get_over_all_file(over_all_info, export_path)
        return True
    except Exception as e:
        print('合同审查出现错误{}'.format(e))
        get_over_all_file([], export_path)
        return False


if __name__ == "__main__":
    print(f"===>PID:{os.getpid()}")
    CORS(app, supports_credentials=True)
    app.run(port=9001)
