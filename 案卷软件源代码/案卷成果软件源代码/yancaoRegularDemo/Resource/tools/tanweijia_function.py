import os

import win32com
from win32com.client import Dispatch
import yancaoRegularDemo.Resource.tools.tangyuhao_function as tyh
from yancaoRegularDemo.Resource.ReadFile import DocxData

"""
为满足新需求，在文件中添加关于涉案金额、罚款比例、罚款金额的批注
一次只打一个同案由其他案件的批注，也就是一个文档的批注
folderPath: 要打批注的文件所在的文件夹 的路径
fileName：要打批注的文件，名称前要加“/”, 不用加后缀名.docx
loc: 要打批注的位置
"""


def addRemarkAboutPenalty(folderPath, fileName, loc, content):
    mw = win32com.client.Dispatch("Word.Application")
    doc = mw.Documents.Open(folderPath + fileName + ".docx")
    tyh.addRemarkInDoc(mw, doc, loc, content)


"""
    获取打批注的内容
"""


def getRemarkContent(referenceFilePath, largest_Dict, cause_of_action, penalty_ratio_range, k, i):
    # folderPath_to_Remark
    remark_list = []
    data = DocxData(referenceFilePath + "/案件处理审批表_.docx")
    # 先获取其他案件的立案编号
    case_Number = data.tabels_content["立案编号"]
    # 再获取同案由其他案件的 涉案金额、罚款比例和罚款金额
    num_list = largest_Dict[cause_of_action][referenceFilePath]
    # 组成批注信息list
    remark_list.append(case_Number)  # remark_list[0]
    remark_list.append(referenceFilePath)  # remark_list[1]
    remark_list.append(num_list[0])  # remark_list[2]
    remark_list.append(num_list[1])  # remark_list[3]
    remark_list.append(num_list[2])  # remark_list[4]
    if penalty_ratio_range[k][i] == 0:  # remark_list[5]
        remark_list.append("在文件中无注明")
    else:
        info = "为" + penalty_ratio_range[k][i][0] + "%-" + penalty_ratio_range[k][i][1] + "%"
        remark_list.append(info)
    return remark_list


# if __name__ == '__main__':

# 避开因为是子串而产生的文件存在检测错误，如想检查“行政处罚决定书_”，但是被“当场行政处罚决定书_”遮蔽
# 例如 is_exist_cover("行政处罚决定书", "当场行政处罚决定书", my_prefix)
def is_exist_cover(detect_name, shadow_name, folder_root):
    file_name = detect_name
    final_file = ""
    for root, dirs, files in os.walk(folder_root):
        for f in files:
            #print(f)
            if file_name + '.docx' == f or file_name + '.doc' == f:
                final_file = f
            elif file_name in f and shadow_name not in f:
                final_file = f
    if final_file == "":
        return False
        # table_father.display(self, "× 《行政处罚决定书》.docx不存在", "red")
    else:
        return final_file
