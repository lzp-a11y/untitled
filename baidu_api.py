import requests
import base64
import json
import urllib3
urllib3.disable_warnings()


def token():
    # 获取token
    url = "https://aip.baidubce.com/oauth/2.0/token"
    grant_type = "client_credentials"
    client_id = "s77qQ5X7iellkGCqFG781nOe"  # api key
    client_secret = "mZr4lpU9xudnxZlL70Mo4lg9uDMnwpcN"  # Secret Key
    data = {"grant_type": grant_type, "client_id": client_id, "client_secret": client_secret}
    res = requests.post(url=url, data=data, verify=False)
    res = res.json()
    # print(json.dumps(res, ensure_ascii=False,indent=2))
    access_token = res["access_token"]
    return access_token


def Verification_code():
    # url = "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic"    # 通用文字识别
    url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"     # 通用文字识别高精度
    header = {"Content-Type":"application/x-www-form-urlencoded"}
    # 二进制方式打开图文件
    f = open(r'E:\自动化测试\MF_Mail\datas\images\pictures2.png', 'rb')
    image = base64.b64encode(f.read())  # 将图片进行base64编码
    access_token = token()
    access_token = access_token  # 30天有效期
    data = {"image":image, "access_token":access_token}
    res2 = requests.post(url=url, data=data, headers=header, verify=False)
    res2 = res2.json()
    # print(res2)
    # print(json.dumps(res2,ensure_ascii=False,indent=2))
    words_result = res2["words_result"]     # 将words_result字段的值取出来
    # print(words_result)        # words_result字段的值是列表
    for a in words_result:       # 遍历列表
        words = a["words"]
        return words


if __name__ == '__main__':
    access_token = Verification_code()
    print(access_token)