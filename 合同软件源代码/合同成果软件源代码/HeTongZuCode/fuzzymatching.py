#!/usr/bin/env python
# -*- encoding: UTF-8 -*-
"""
@Project: ChinaTobaccoContract_V2
@File: fuzzymatching.py
@Author: Mr.Blonde
@Date: 2021/12/13 10:23
"""
# !/usr/bin/python
# encoding=utf-8
import numpy as np
from fuzzywuzzy import fuzz
import torch
from sentence_transformers.util import cos_sim
from text2vec import SBert

# import os
# os.environ["CUDA_VISIBLE_DEVICES"] = "0"

print("加载Sbert………………")
device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
# device = torch.device("cpu")
print("运行设备：", device)
model = SBert('paraphrase-multilingual-MiniLM-L12-v2')
# torch.save(model,"./sbert.pkl")
# model = torch.load("./sbert.pkl").to(device)
print("加载Sbert完成…………")


def text2vecComputeScores(list_items, paragraphs):
    embeddings1 = model.encode(list_items)
    embeddings2 = model.encode(paragraphs)
    scores = cos_sim(embeddings1, embeddings2).to('cpu').numpy()
    return scores


# 通过不同的方式算出str1和str2的相似度,基于编辑距离Levenshtein Distance, 这个相似度标准化到0～100
def fuzzyMatching(str1s, str2s, method="Simple Ratio"):
    scores = np.zeros((len(str1s), len(str2s)))
    if method == "text2vec":
        scores = text2vecComputeScores(str1s, str2s) * 100
    else:
        for idx1, str1 in enumerate(str1s):
            for idx2, str2 in enumerate(str2s):
                str1 = fullToHalf(str1).replace(' ', '')
                str2 = fullToHalf(str2).replace(' ', '')
                score = 0
                if method == "Simple Ratio":
                    score = fuzz.ratio(str1, str2)
                elif method == "Partial Ratio":
                    score = fuzz.partial_ratio(str1, str2)
                elif method == "Token Sort Ratio":
                    score = fuzz.token_set_ratio(str1, str2)
                elif method == "Token Set Ratio":
                    score = fuzz.token_set_ratio(str1, str2)
                scores[idx1, idx2] = score
    return scores


def extract_items(items, paragraphs, scores, vote, vote_type, vote_num=2):
    rss = {}
    item_index = 0
    if vote > 0 and vote_type == "standard":
        shape = scores.shape
        for key in items.keys():
            b = []
            for i, key1 in enumerate(items[key].keys()):
                collection = {}
                for j in range(shape[0]):
                    one = scores[j][item_index]
                    sorted_idx = np.argsort(-one).tolist()[0:vote]
                    for k in sorted_idx:
                        if paragraphs[k] in collection.keys():
                            collection[paragraphs[k]] += 1
                        else:
                            collection[paragraphs[k]] = 1
                # 字典降序
                collection = sorted(collection.items(), key=lambda kv: kv[1], reverse=True)
                a = []
                for k in collection:
                    key_ = k[0]
                    value_ = k[1]
                    if value_ >= vote_num:
                        a.append([key1, key_, value_])
                b.append(a)
                item_index += 1
            rss[key] = [b]

        return rss
    for key in items.keys():
        b = []
        for i, key1 in enumerate(items[key].keys()):
            thr = items[key][key1]
            if vote > 0 and vote_type == "mean":
                thr = vote_num
            type_ = str(type(thr)).split(" ")[1].replace('>', '').replace('\'', '')
            score = scores[item_index]
            sorted_index = np.argsort(-score)
            if type_ == 'int':
                k1_index = sorted_index[0:thr].tolist()
                a = []
                for j in k1_index:
                    a.append([key1, paragraphs[j], score[j]])
                b.append(a)
            elif type_ == "float":
                if thr <= 1.0:
                    thr *= 100
                one = score[sorted_index].tolist()
                f_i = 0
                for one_i, s in enumerate(one):
                    f_i = one_i
                    if s < thr:
                        break

                k1_index = sorted_index[0:f_i].tolist()
                a = []
                for j in k1_index:
                    # if len(paragraphs[j]) >=2:
                    a.append([key1, paragraphs[j], score[j]])
                b.append(a)
            item_index += 1
        rss[key] = [b]
    return rss


# 全角半角转换表
definedConverts = {
    "　": "",
    "“": '',
    "”": '',
    "！": "",
    "￥": "",
    "……": "",
    "（": "",
    "）": "",
    "——": "",
    "【": "",
    "】": "",
    "；": ";",
    "’": "'",
    "：": ":",
    "|": "",
    "、": "",
    "？": "?",
    ".": "",
    "\t": "",
}


# 全角转半角
def fullToHalf(str_):
    for from_, to_ in definedConverts.items():
        # print(from_, to_)
        str_ = str_.replace(from_, to_)
    return str_


# 半角转全角
def halfToFull(str_):
    for from_, to_ in definedConverts.items():
        # print(from_, to_)
        str_ = str_.replace(to_, from_)
    return str_


# 买卖合同批注处理
def maimaihetong():
    1
