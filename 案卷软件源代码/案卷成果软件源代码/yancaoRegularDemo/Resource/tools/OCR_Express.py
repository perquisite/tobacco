import requests
import base64

"""
快递面单识别
要求base64编码和urlencode后大小不超过4M，最短边至少15px，最长边最大4096px，支持jpg/jpeg/png/bmp格式
"""


class OCR_Express:
    request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/waybill"
    access_token = '24.decf5cac15a1d244e14344531c1e8807.2592000.1645931028.282335-25565290'
    request_url = request_url + "?access_token=" + access_token
    headers = {'content-type': 'application/x-www-form-urlencoded'}

    def __init__(self, img_path):
        # 二进制方式打开图片文件
        self.img_path = img_path
        f = open(self.img_path, 'rb')
        self.img = base64.b64encode(f.read())
        self.params = {"image": self.img}
        self.request_url = OCR_Express.request_url + "?access_token=" + OCR_Express.access_token
        self.headers = {'content-type': 'application/x-www-form-urlencoded'}

        response = requests.post(self.request_url, data=self.params, headers=self.headers)
        if response:
            self.result = response.json()
        else:
            print("No response")

    # 获取 寄件人姓名
    def getSenderName(self):
        return self.result['words_result'][0]['sender_name'][0]['word']

    # 获取 寄件人地址
    def getSenderAddr(self):
        return self.result['words_result'][0]['sender_addr'][0]['word']

    # 获取 寄件人电话
    def getSenderPhone(self):
        return self.result['words_result'][0]['sender_phone'][0]['word']

    # 获取 收件人姓名
    def getRecipientName(self):
        return self.result['words_result'][0]['recipient_name'][0]['word']

    # 获取 收件人地址
    def getRecipientAddr(self):
        return self.result['words_result'][0]['recipient_addr'][0]['word']

    # 获取 收件人电话
    def getRecipientPhone(self):
        return self.result['words_result'][0]['recipient_phone'][0]['word']

    # 获取 快递运单号
    def getWaybillNumber(self):
        return self.result['words_result'][0]['waybill_number'][0]['word']

    # 获取 三段码
    def get3SegmentCode(self):
        return self.result['words_result'][0]['three_segment_code'][0]['word']

    # 获取 条形码
    def getBarCode(self):
        return self.result['words_result'][0]['bar_code'][0]['word']


if __name__ == '__main__':
    sb1 = OCR_Express(r'C:\Users\twj\Desktop\1.jpg')
    sb2 = OCR_Express(r'C:\Users\twj\Desktop\2.jpg')
    sb3 = OCR_Express(r'C:\Users\twj\Desktop\3.jpg')
    sb4 = OCR_Express(r'C:\Users\twj\Desktop\4.jpg')
    sb5 = OCR_Express(r'C:\Users\twj\Desktop\5.jpg')
    print(sb1.result)
    print(sb2.result)
    print(sb3.result)
    print(sb4.result)
    print(sb5.result)
    print("\n")
    print(sb1.getSenderAddr())
    print(sb1.getSenderName())
    print(sb1.getSenderPhone())
    print(sb1.getSenderAddr())
    print(sb1.getSenderName())
    print(sb1.getSenderPhone())
    print(sb1.getWaybillNumber())
    print(sb1.get3SegmentCode())
    print(sb1.getBarCode())
    print("\n")
    print(sb2.getSenderAddr())
    print(sb2.getSenderName())
    print(sb2.getSenderPhone())
    print(sb2.getSenderAddr())
    print(sb2.getSenderName())
    print(sb2.getSenderPhone())
    print(sb2.getWaybillNumber())
    print(sb2.get3SegmentCode())
    print(sb2.getBarCode())
    print("\n")
    print(sb3.getSenderAddr())
    print(sb3.getSenderName())
    print(sb3.getSenderPhone())
    print(sb3.getSenderAddr())
    print(sb3.getSenderName())
    print(sb3.getSenderPhone())
    print(sb3.getWaybillNumber())
    print(sb3.get3SegmentCode())
    print(sb3.getBarCode())
    print("\n")
    print(sb4.getSenderAddr())
    print(sb4.getSenderName())
    print(sb4.getSenderPhone())
    print(sb4.getSenderAddr())
    print(sb4.getSenderName())
    print(sb4.getSenderPhone())
    print(sb4.getWaybillNumber())
    print(sb4.get3SegmentCode())
    print(sb4.getBarCode())
    print("\n")
    print(sb5.getSenderAddr())
    print(sb5.getSenderName())
    print(sb5.getSenderPhone())
    print(sb5.getSenderAddr())
    print(sb5.getSenderName())
    print(sb5.getSenderPhone())
    print(sb5.getWaybillNumber())
    print(sb5.get3SegmentCode())
    print(sb5.getBarCode())
