# coding:utf-8
"""
    Project:中国烟草案卷执法组文本实体识别类
    Author:谢俊宇
    Date:2021-12-06 9:30
"""
"""
标签  含义	   标签	  含义	   标签	  含义	   标签	 含义
 n	普通名词	    f	方位名词	    s	处所名词	    nw	作品名
 nz	其他专名	    v	普通动词	    vd	动副词	    vn	名动词
 a	形容词	    ad	副形词	    an	名形词	    d	副词
 m	数量词	    q	量词	        r	代词	        p	介词
 c	连词	        u	助词	        xc	其他虚词	    w	标点符号
 PER 人名	    LOC	地名	        ORG	机构名	    TIME 时间
"""
from LAC import LAC


class EntityRecognition(object):
    def __init__(self):
        super(EntityRecognition, self).__init__()

    """
        获取句子分词结果
    """

    def get_seg_result(self, input_text):
        lac = LAC(mode='seg')
        return lac.run(input_text)

    """
        获取词性标注与实体识别结果
    """

    def get_ner_result(self, input_text):
        lac = LAC(mode='lac')
        return lac.run(input_text)

    """
        返回重要程度
    """

    def get_importancef_rank(self, input_text):
        lac = LAC(mode='rank')
        return lac.run(input_text)

    def get_identity_with_tag(self, input_text, tag):
        result_list = self.get_ner_result(input_text)
        tag_list = [i for i, v in enumerate(result_list[1]) if v == tag]
        return [result_list[0][i[1]] for i in enumerate(tag_list)]


if __name__ == '__main__':
    text = u"违法事实：2021年4月25日，当事人程孝军在其经营门市从一名中年男子处以210.00元的价格购买了84mm娇子（五粮醇香）1条、以235.00元的价格购买了97mm南京（炫赫门炫彩）1条、以300.00元每条的价格购买了84mm钻石（荷花） 2条、以310.00元每条价格购买了84mm娇子（五粮浓香细支）1条和84mm娇子（软宽窄平安）2条、以400.00元的价格购买了84mm中华（ 硬）1条、以390.00元的价格购买了88mm天子（中国心中支）1条，共计7个品种9条卷烟，共付烟款2760.00元，后将所购卷烟放在其经营门市准备销售。"
    entityrecognition = EntityRecognition()
    print(entityrecognition.get_seg_result(text))
    print(entityrecognition.get_ner_result(text))
    print(entityrecognition.get_identity_with_tag(text, 'LOC'))
    print(entityrecognition.get_importancef_rank(text))

    # texts = [u"LAC是个优秀的分词工具", u"百度是一家高科技公司"]
    # print(entityrecognition.get_seg_result(texts))
    # print(entityrecognition.get_ner_result(texts))
    # print(entityrecognition.get_importancef_rank(texts))
