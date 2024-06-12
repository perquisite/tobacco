# -*- coding:utf-8 -*-
# @ModuleName: ContractInfo
# @Function: 
# @Author: huhonghui
# @email: 1241328737@qq.com
# @Time: 2021/6/18 16:59

class ContractCheckResult:
    def __init__(self, contract_type, factors_dic, factors_ok_list, factors_error_list, factors_to_inform_dic,
                 type_name=None):
        self.type = contract_type
        self.type_name = type_name
        self.factors = factors_dic
        self.factors_ok = factors_ok_list
        self.factors_error = factors_error_list
        self.factors_to_inform = factors_to_inform_dic
