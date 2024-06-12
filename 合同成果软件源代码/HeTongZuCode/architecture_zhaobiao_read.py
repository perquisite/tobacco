import re
from docx import Document

class A_DocReader:
    def __init__(self, file_path):
        self.file_path = file_path

    def get_text(self):
        try:
            document = Document(self.file_path)
        except:
            print("请确定\"" + self.file_path + "\"不是空文档！")  ##################################
            return "docx_blank"
        paragraghs = document.paragraphs
        text = ""
        for p in paragraghs:
            if p.text != "":
                # 把半角全角符号一律转全角 add by qy
                text = text.replace(':', '：')
                text = text.replace('(', '（')
                text = text.replace(')', '）')
                text = text.replace('\ue5e5', ' ').replace('\u3000', ' ')
                text += p.text + "\n"
        return text

    def get_info(self,text):
        result_dict={}
        try:
            match = '签约合同价暂定为：\n人民币¥(.*)（大写）；\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['签约合同价']=factor
        except:
            result_dict['签约合同价']=None

        try:
            match = '安全文明施工费：\n人民币（大写）.*（¥(.*)元）；\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['安全文明施工费']=factor
        except:
            result_dict['安全文明施工费']=None

        try:
            match = '.*材料和工程设备暂估价金额：\n人民币（大写）.*（¥(.*)元）；\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['材料和工程设备暂估价金额'] = factor
        except:
            result_dict['材料和工程设备暂估价金额'] = None

        try:
            match = '.*专业工程暂估价金额：\n人民币（大写）.*（¥(.*)元）；\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['专业工程暂估价金额'] = factor
        except:
            result_dict['专业工程暂估价金额'] = None


        try:
            match = '.*暂列金额：\n人民币（大写）.*（¥(.*)元）。\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['暂列金额'] = factor
        except:
            result_dict['暂列金额'] = None

        try:
            match = '工程内容：(.*)\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "")
            result_dict['工程内容'] = factor
        except:
            result_dict['工程内容'] = None

        try:
            match = '承包人项目经理：(.*)\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "")
            result_dict['承包人项目经理'] = factor
        except:
            result_dict['承包人项目经理'] = None

        try:
            match = '工程名称：(.*)\n工程地点'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['工程名称']=factor
        except:
            result_dict['工程名称']=None

        try:
            match = '工程地点：(.*)\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['工程地点']=factor
        except:
            result_dict['工程地点']=None

        try:
            match = '工程立项批准文号：(.*)\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['工程立项批准文号']=factor
        except:
            result_dict['工程立项批准文号']=None

        try:
            match = '资金来源：(.*)\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['资金来源']=factor
        except:
            result_dict['资金来源']=None

        try:
            match = '工程内容：(.*)\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['工程内容']=factor
        except:
            result_dict['工程内容']=None

        try:
            match = '工程内容：(.*)\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['工程内容']=factor
        except:
            result_dict['工程内容']=None

        try:
            match = '计划工期：(.*)个日历天.*\n'
            factor = re.findall(match, text)[0].replace(" ", "").replace("；", "").replace("。", "").replace("\ue5e5", "")
            result_dict['计划工期']=factor
        except:
            result_dict['计划工期']=None


        return result_dict


# filePath = r'C:\Users\12259\Desktop\招标文件.docx'
# reader=A_DocReader(filePath)
# text=reader.get_text()
# architecture_dict=reader.get_info(text)
# for k,v in architecture_dict.items():
#     print(k)
#     print(v)
