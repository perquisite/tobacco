import os
import re

import cn2an
import rdflib

# from yancaoRegularDemo.Resource.table_19_36.Table34 import *
from yancaoRegularDemo.Resource.tools.utils import get_root_dir
from zhishitupu.src.function import getRootPath
from zhishitupu.tools.discretionary_function import *

"""
info_one:
[{'案由':'未在当地烟草专卖批发企业进货', '处罚结果':['处以xxx。','予以xxx。']}, {'案由':'未在当地烟草专卖批发企业进货', '处罚结果':['处以xxx。','予以xxx。']}]
info_two:
{'从轻':True or False,'存在从轻证据':['xxxxxx','xxxx'],'从重':True or False}
"""


def discretionary_power(info_one, info_two, city='成都市局'):
    return_tip = ""
    light_heavy_info_to_print = about_light_heavy(info_two)
    return_tip += light_heavy_info_to_print
    # return light_heavy_info
    info_one_processed = []  # 该list的每个元素形如{案由：xx, 理应当次： ，实际档次：应该没收：T/F， 实际没收：T/F，是否罚款：T/F，理应条文：‘’}的字典，用作输出依据
    return_tip += light_heavy_info_to_print
    for every_case_info in info_one:
        every_tip = ""
        written_reason = every_case_info['案由']
        punishment = every_case_info['处罚结果']
        real_reason = getRealReason(written_reason)
        info_processed = {}  # 是info_one_processed（list）的一个元素
        if real_reason is None:  # 没有找到 真实案由
            return_tip += "案由“" + written_reason + "”不合规或没有找到该案由，请人工检查！"
        else:
            info_processed = get_every_detail(written_reason, real_reason, punishment, city)
            # info_processed = get_every_detail('为存储假冒伪劣烟草制品、走私烟草制品的烟草制品提供条件', '为生产、运输、邮寄、存储、销售假冒伪劣烟草制品、走私烟草制品、出口倒流国产烟草制品、未缴付关税而流出免税店和保税区的烟草制品提供条件',
            #                                   ['处以物品价值20000元 10000元的罚款。', '予以没收。'], city)
            # info_processed = get_every_detail('无烟草专卖批发企业许可证经营烟草制品批发业务',
            #                                  '无烟草专卖批发企业许可证经营烟草制品批发业务', ['将250条卷烟予以没收。'], city)
            # info_one = [{'案由': '为存储假冒伪劣烟草制品、走私烟草制品的烟草制品提供条件', '处罚结果': ['处以物品价值2500元 5010元的罚款。', '予以没收、收购。']}]
            # print(info_processed)
            if info_processed:
                # every_tip += "案由“" + info_processed['案由'] + "”的处罚结果为——"
                # every_tip += "“" + str(punishment) + "”" + ","
                every_tip += "案由“" + info_processed['案由'] + "”的处罚结果"
                if str(punishment) == '[]':
                    every_tip += "由于书写规范问题未成功提取，请人工检验自由裁量权！"
                else:
                    every_tip += "为——“" + str(punishment) + "”" + ","
                shiji = info_processed['实际档次']
                liying = info_processed['理应档次']
                if liying is None:  # 理应档次 无法确定
                    every_tip += '无法确定理应的处罚档次，请人工核查！'
                else:
                    if shiji is not None:  # 理应档次和实际档次 都确定
                        every_tip += "属于该地自由裁量基准尺度第" + str(shiji) + "项，"
                        every_tip += "其基准尺度为第" + str(liying) + "项，"
                        if shiji - liying < 0:
                            every_tip += "属于从轻处罚"
                        elif shiji - liying > 0:
                            every_tip += "属于从重处罚"
                    else:  # 理应档次确定 实际档次没有
                        every_tip += "但无法确定实际的处罚档次，理应档次对应的法律条文为：“" + info_processed['理应条文'] + "”,且请人工核查！"
                if info_processed['应该没收'] is True and info_processed['实际没收'] is False:
                    every_tip += "。理应予以没收，但实际并没有没收。"
                elif info_processed['应该没收'] is False and info_processed['实际没收'] is True:
                    every_tip += "。不应予以没收，但实际予以了没收。"
            else:
                every_tip = "由于案由撰写不规范，所以在库中未提取到对应案由的信息！"

        return_tip += every_tip
    return return_tip
    # return return_tip


if __name__ == "__main__":
    # info_one = [{'案由': '销售、运输、存储、投递走私烟草制品，出口倒流国产烟草制品、未缴付关税而流出免税店和保税区的烟草制品', '处罚结果': ['处以60001元 3000元的罚款。']}]
    # '处以1001元 6%的罚款。',
    # info_one = [{'案由': '为存储假冒伪劣烟草制品、走私烟草制品的烟草制品提供条件', '处罚结果': ['处以物品价值2500元 5010元的罚款。', '予以没收、收购。']}]
    info_one = [{'案由': '无烟草专卖批发企业许可证经营烟草制品批发业务', '处罚结果': ['将250条卷烟予以没收。']}]
    #info_one = get_every_detail('无烟草专卖批发企业许可证经营烟草制品批发业务', '无烟草专卖批发企业许可证经营烟草制品批发业务', ['将250条卷烟予以没收。'],)
    info_two = {'从轻': True, '存在从轻证据': ['xxxxxx', 'yyyy'], '从重': False}
    print(discretionary_power(info_one, info_two, city='成都市局'))
