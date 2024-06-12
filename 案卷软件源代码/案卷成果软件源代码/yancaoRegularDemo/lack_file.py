import os

list_right = ['立案（不予立案）报告表', '延长立案审批表', '延长立案告知书', '指定管辖通知书', '检查（勘验）笔录',  '证据先行登记保存通知书', '抽样取证物品清单',
              '询问笔录',
              '涉案烟草专卖品核价表', '证据复制（提取）单', '公告', '损耗费用审批表', '案件移送函', '案件移送回执', '移送财物清单', '协助调查函', '撤销立案报告表', '案件调查终结报告',
              '延长调查终结审批表', '延长调查期限告知书', '先行登记保存证据处理通知书', '涉案物品返还清单', '行政处罚事先告知书', '陈述申辩记录', '听证告知书', '听证通知书',
              '不予受理听证通知书',
              '听证公告', '听证笔录', '听证报告', '案件集体讨论记录', '案件处理审批表', '当场行政处罚决定书', '行政处罚决定书', '行政处理决定书', '不予行政处罚决定书', '送达回证',
              '送达公告',
              '责令改正通知书', '责令整顿通知书', '整顿终结通知书', '违法物品销毁记录表', '罚没变价处理审批表', '罚没物品移交单', '加处罚款决定书', '强制执行申请书', '延期缴款审批表',
              '对协助办案有功单位、个人授奖呈报表', '结案报告表', '卷宗封面', '卷宗目录', '卷内备考表', '关于同意延长证据登记保存期限的批复', '关于延长证据登记保存期限的请示',
              '延长案件调查终结审批表_', '证据先行登记保存批准书_']


def match(l, l0):
    if l in l0 or l0 in l:
        return True
    else:
        return False


def lack_file(list):
    list_lack = []
    for l0 in list_right:
        flag = 0
        for l in list:
            if match(l, l0):
                flag = 1
                break
        if flag == 0:
            list_lack.append(l0)
    return list_lack


def lack_file_dir(dir):
    for root, dirs, files in os.walk(dir):
        return lack_file(files)


if __name__ == "__main__":
    print(lack_file_dir(r"C:\Users\Zero\Desktop\烟草文书demo\2021184117_崇烟立2021第1号"))
    print(len(lack_file_dir(r"C:\Users\Zero\Desktop\烟草文书demo\2021184117_崇烟立2021第1号")))
    print(len(list_right))
