# -*- coding:utf-8 -*-
# @ModuleName: ContractType
# @Function: 
# @Author: huhonghui
# @email: 1241328737@qq.com
# @Time: 2021/6/18 16:38

from enum import Enum


class ContractType(Enum):
    NotSure = 0
    BuySellContract = 1
    RentContract = 2
    ConstructionContract = 3
    PurchaseAndWarehousingContract = 4
    PropertyManagementContract = 5

    # 非标准买卖合同
    BuySellContract_Not_Standard = 6
