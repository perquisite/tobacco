import rdflib
import json
import pandas as pd
import os
import numpy as np
import cn2an


def excle_to_rdf(file, save_path):
    data = pd.read_excel(file)
    try:
        if not np.any(data.isnull()):
            pass
    except:
        print('表格中有空未填写')
        return 0

    label = data.columns.values
    json_dic = {}
    for item in data['条目']:
        index = data[data['条目'] == '违法行为'].index.tolist()[0]
        if item == '违法行为':
            for l in label[1:]:
                value = data[l].loc[index]
                if value != '无':
                    json_dic.update({l:value})
        elif '处罚结果' in item:
            index0 = data[data['条目'] == item].index.tolist()[0]
            item_dic = {}
            for l0 in label[1:]:
                value0 = data[l0].loc[index0]
                if value0 != '无':
                    if value0 == '是':
                        value0 = 'true'
                    if value0 == '否':
                        value0 = 'false'
                    item_dic.update({l0:value0})
            json_dic.update({item:item_dic})
        else:
            print('有不识别字符')

    keys = list(json_dic.keys())
    g = rdflib.Graph()
    s = rdflib.URIRef('http://www.tobacco.org#违法行为' + str(cn2an.cn2an(json_dic['法律条数'])))
    for key in keys:
        if '处罚结果' in key:
            p1 = rdflib.URIRef('http://www.tobacco.org#处罚')
            o1 = rdflib.URIRef('http://www.tobacco.org#违法行为' + str(cn2an.cn2an(json_dic['法律条数'])) + key)
            g.add((s, p1, o1))

            keys1 = list(json_dic[key].keys())
            for key1 in keys1:
                p3 = rdflib.URIRef('http://www.tobacco.org#' + key1)
                o3 = rdflib.term.Literal(json_dic[key][key1])
                g.add((o1, p3, o3))
        else:
            p2 = rdflib.URIRef('http://www.tobacco.org#' + key)
            o2 = rdflib.term.Literal(json_dic[key])
            g.add((s, p2, o2))
    g.serialize(save_path, format="xml")
    print('已按照《' + json_dic['法律名称'] + '》第' + json_dic['法律条数'] + '条生成对应法条知识图谱（RDF）')

# def excle_to_rdf(file, save_path):
#     data = pd.read_excel(file)
#     try:
#         if not np.any(data.isnull()):
#             pass
#     except:
#         print('表格中有空未填写')
#         return 0
#
#     label = data.columns.values
#     json_dic = {}
#     for item in data['条目']:
#         index = data[data['条目'] == '违法行为'].index.tolist()[0]
#         if item == '违法行为':
#             for l in label[1:]:
#                 value = data[l].loc[index]
#                 if value != '无':
#                     json_dic.update({l:value})
#         elif '处罚结果' in item:
#             index0 = data[data['条目'] == item].index.tolist()[0]
#             item_dic = {}
#             for l0 in label[1:]:
#                 value0 = data[l0].loc[index0]
#                 if value0 != '无':
#                     if value0 == '是':
#                         value0 = 'true'
#                     if value0 == '否':
#                         value0 = 'false'
#                     item_dic.update({l0:value0})
#             json_dic.update({item:item_dic})
#         else:
#             print('有不识别字符')
#
#     keys = list(json_dic.keys())
#     g = rdflib.Graph()
#     s = rdflib.URIRef('http://www.tobacco.org#' + json_dic['法律名称'] + json_dic['法律条数'])
#     for key in keys:
#         if '处罚结果' in key:
#             p1 = rdflib.URIRef('http://www.tobacco.org#处罚')
#             o1 = rdflib.URIRef('http://www.tobacco.org#' + key)
#             g.add((s, p1, o1))
#
#             keys1 = list(json_dic[key].keys())
#             for key1 in keys1:
#                 p3 = rdflib.URIRef('http://www.tobacco.org#' + key1)
#                 o3 = rdflib.term.Literal(json_dic[key][key1])
#                 g.add((o1, p3, o3))
#         else:
#             p2 = rdflib.URIRef('http://www.tobacco.org#' + key)
#             o2 = rdflib.term.Literal(json_dic[key])
#             g.add((s, p2, o2))
#     g.serialize(save_path, format="xml")
#     print('已按照《' + json_dic['法律名称'] + '》第' + json_dic['法律条数'] + '条生成对应法条知识图谱（RDF）')


if __name__ == "__main__":
    excle_to_rdf(r'C:\Users\twj\Desktop\KnowledgeGraph\test.xlsx', r'C:/Users/twj/Desktop/111.rdf')
