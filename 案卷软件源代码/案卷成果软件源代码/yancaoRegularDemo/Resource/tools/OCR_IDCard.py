import requests
from aip import AipOcr
"""   
身份证信息识别
base64编码后进行urlencode，要求base64编码和urlencode后大小不超过4M，最短边至少15px，最长边最大4096px,支持jpg/jpeg/png/bmp格式
输入：
    参数1：正面图片的路径，参数2：反面图片的路径。事实上，参数1和2可以调换。如果只需要一个参数，则另一个要输入""(空字符串)
输出（都是字符串）：
    1、‘no response’：网络等原因造成的没有响应
    2、‘no image input’：没有检测到图片路径的输入
    3、‘non_idcard’: 输入的图片不是身份证
    3、正常结果字符串
"""

class OCR_IDCard:
    APP_ID = '25565290'
    API_KEY = '6d1G2RDtsODZoN5c46ifWxvD'
    SECRET_KEY = 'AdSejl9OvpEGs9MZY8slrwjCtdiC0n4X'
    options = {}
    options["detect_direction"] = "true"
    options["detect_risk"] = "false"

    def __init__(self, img_path_front, img_path_back):
        self.result = {}
        self.result_front = {}
        self.result_back = {}
        self.flag = False
        self.client = AipOcr(OCR_IDCard.APP_ID, OCR_IDCard.API_KEY, OCR_IDCard.SECRET_KEY)
        # 正面
        if not img_path_front == "":
            idCardSide = "front"
            self.img = OCR_IDCard.get_file_content(img_path_front)
            try:
                response = self.client.idcard(self.img, idCardSide, OCR_IDCard.options)
                if type(response) == dict:
                    self.result_front = response
            except Exception as e:
                #print(type(e))
                if type(e) == requests.exceptions.ConnectionError:
                    self.result_front = 'no response'
            #print(self.result_front)
        else:
            self.result_front = "no image input"
        # 再获取反面
        if not img_path_back == "":
            idCardSide = "back"
            self.img = OCR_IDCard.get_file_content(img_path_back)
            try:
                response = self.client.idcard(self.img, idCardSide, OCR_IDCard.options)
                if type(response) == dict:
                    self.result_back = response
            except Exception as e:
                # print(type(e))
                if type(e) == requests.exceptions.ConnectionError:
                    self.result_back = 'no response'
            # print(self.result_back)
        else:
            self.result_back = "no image input"

        # 合并两次识别的结果
        if type(self.result_back) == dict and type(self.result_front) == dict:
            self.flag = True  # True代表传了两个路径
            res = {**self.result_front["words_result"], **self.result_back["words_result"]}
            self.result['words_result'] = res
            #print("最后合并结果\n")
            #print(self.result)

    # 获取文件信息
    def get_file_content(filePath):
        with open(filePath, 'rb') as fp:
            return fp.read()

    # 获取 姓名
    def getName(self):
        if self.flag == True:
            return self.result['words_result']['姓名']['words']
        else:
            if (not self.result_front in ["no image input", "no response"] and self.result_front['image_status'] == 'reversed_side') and self.result_back in ["no image input", "no response"]:
                return self.result_back
            elif type(self.result_front) == dict and self.result_back in ["no image input", "no response"]:
                return self.result_front['words_result']['姓名']['words']
            elif self.result_front in ["no image input", "no response"] and (not self.result_back in ["no image input", "no response"] and self.result_back['image_status'] == 'reversed_side'):
                return self.result_back['words_result']['姓名']['words']
            elif self.result_front in ["no image input", "no response"] and type(self.result_back) == dict:
                return self.result_front
            else:
                return self.result_front

    # 获取 性别
    def getSex(self):
        if self.flag == True:
            return self.result['words_result']['性别']['words']
        else:
            if (not self.result_front in ["no image input", "no response"] and self.result_front[
                'image_status'] == 'reversed_side') and self.result_back in ["no image input", "no response"]:
                return self.result_back
            elif type(self.result_front) == dict and self.result_back in ["no image input", "no response"]:
                return self.result_front['words_result']['性别']['words']
            elif self.result_front in ["no image input", "no response"] and (
                    not self.result_back in ["no image input", "no response"] and self.result_back[
                'image_status'] == 'reversed_side'):
                return self.result_back['words_result']['性别']['words']
            elif self.result_front in ["no image input", "no response"] and type(self.result_back) == dict:
                return self.result_front
            else:
                return self.result_front

    # 获取 民族
    def getNation(self):
        if self.flag == True:
            return self.result['words_result']['民族']['words']
        else:
            if (not self.result_front in ["no image input", "no response"] and self.result_front[
                'image_status'] == 'reversed_side') and self.result_back in ["no image input", "no response"]:
                return self.result_back
            elif type(self.result_front) == dict and self.result_back in ["no image input", "no response"]:
                return self.result_front['words_result']['民族']['words']
            elif self.result_front in ["no image input", "no response"] and (
                    not self.result_back in ["no image input", "no response"] and self.result_back[
                'image_status'] == 'reversed_side'):
                return self.result_back['words_result']['民族']['words']
            elif self.result_front in ["no image input", "no response"] and type(self.result_back) == dict:
                return self.result_front
            else:
                return self.result_front

    # 获取 出生日期
    def getBirth(self):
        if self.flag == True:
            return self.result['words_result']['出生']['words']
        else:
            if (not self.result_front in ["no image input", "no response"] and self.result_front[
                'image_status'] == 'reversed_side') and self.result_back in ["no image input", "no response"]:
                return self.result_back
            elif type(self.result_front) == dict and self.result_back in ["no image input", "no response"]:
                return self.result_front['words_result']['出生']['words']
            elif self.result_front in ["no image input", "no response"] and (
                    not self.result_back in ["no image input", "no response"] and self.result_back[
                'image_status'] == 'reversed_side'):
                return self.result_back['words_result']['出生']['words']
            elif self.result_front in ["no image input", "no response"] and type(self.result_back) == dict:
                return self.result_front
            else:
                return self.result_front

    # 获取 身份号码
    def getIDnumber(self):
        if self.flag == True:
            return self.result['words_result']['公民身份号码']['words']
        else:
            if (not self.result_front in ["no image input", "no response"] and self.result_front['image_status'] == 'reversed_side') and self.result_back in ["no image input", "no response"]:
                return self.result_back
            elif type(self.result_front) == dict and self.result_back in ["no image input", "no response"]:
                return self.result_front['words_result']['公民身份号码']['words']
            elif self.result_front in ["no image input", "no response"] and (not self.result_back in ["no image input", "no response"] and self.result_back['image_status'] == 'reversed_side'):
                return self.result_back['words_result']['公民身份号码']['words']
            elif self.result_front in ["no image input", "no response"] and type(self.result_back) == dict:
                return self.result_front
            else:
                return self.result_front

    # 获取 住址
    def getAddress(self):
        if self.flag == True:
            return self.result['words_result']['住址']['words']
        else:
            if (not self.result_front in ["no image input", "no response"] and self.result_front['image_status'] == 'reversed_side') and self.result_back in ["no image input", "no response"]:
                return self.result_back
            elif type(self.result_front) == dict and self.result_back in ["no image input", "no response"]:
                return self.result_front['words_result']['住址']['words']
            elif self.result_front in ["no image input", "no response"] and (not self.result_back in ["no image input", "no response"] and self.result_back['image_status'] == 'reversed_side'):
                return self.result_back['words_result']['住址']['words']
            elif self.result_front in ["no image input", "no response"] and type(self.result_back) == dict:
                return self.result_front
            else:
                return self.result_front

    # 获取 失效日期
    def getExpiringDate(self):
        if self.flag == True:
            return self.result['words_result']['失效日期']['words']
        else:
            if (not self.result_front in ["no image input", "no response"] and self.result_front['image_status'] == 'reversed_side') and self.result_back in ["no image input", "no response"]:
                return self.result_front['words_result']['失效日期']['words']
            elif type(self.result_front) == dict and self.result_back in ["no image input", "no response"]:
                return self.result_back
            elif self.result_front in ["no image input", "no response"] and (not self.result_back in ["no image input", "no response"] and self.result_back['image_status'] == 'reversed_side'):
                return self.result_front
            elif self.result_front in ["no image input", "no response"] and type(self.result_back) == dict:
                return self.result_back['words_result']['失效日期']['words']
            else:
                return self.result_back



if __name__ == '__main__':

    #sb1 = OCR_IDCard(r'C:\Users\twj\Desktop\2-1.jpg', r'C:\Users\twj\Desktop\1-1.jpg')
    #print(sb1.result_front)
    #print(sb1.result_back)
    # print(sb1.getName())
    # print(sb1.getExpiringDate())
    # print(sb1.getIDnumber())
    # print(sb1.getAddress())
    # print(sb1.getBirth())
    # print(sb1.getNation())
    # print(sb1.getSex())
    # print(sb1.getExpiringDate())
    # print("\n")
    sb2 = OCR_IDCard('', r'C:\Users\twj\Desktop\1-1.jpg')
    print(sb2.getName())
    print(sb2.getIDnumber())
    print(sb2.getAddress())
    print(sb2.getBirth())
    print(sb2.getNation())
    print(sb2.getSex())
    print(sb2.getExpiringDate())
    # print("\n")
    # sb3 = OCR_IDCard(r'C:\Users\twj\Desktop\2-1.jpg', '')
    # print(sb3.getName())
    # print(sb3.getIDnumber())
    # print(sb3.getAddress())
    # print(sb3.getBirth())
    # print(sb3.getNation())
    # print(sb3.getSex())
    # print(sb3.getExpiringDate())
    # print("\n")
    # sb4 = OCR_IDCard(r'C:\Users\twj\Desktop\1-1.jpg', r'C:\Users\twj\Desktop\2-1.jpg')
    # print(sb4.getName())
    # print(sb4.getIDnumber())
    # print(sb4.getAddress())
    # print(sb4.getBirth())
    # print(sb4.getNation())
    # print(sb4.getSex())
    # print(sb4.getExpiringDate())
    # print("\n")
    # sb5 = OCR_IDCard('', '')
    # print(sb5.getName())
    # print(sb5.getIDnumber())
    # print(sb5.getAddress())
    # print(sb5.getBirth())
    # print(sb5.getNation())
    # print(sb5.getSex())
    # print(sb5.getExpiringDate())

