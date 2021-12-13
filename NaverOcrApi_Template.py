import json
import base64
import requests

with open("체크박스.JPG", "rb") as f:
    img = base64.b64encode(f.read())

 # 본인의 APIGW Invoke URL로 치환
URL = "https://0e55e899676d4879857d309ed64c2688.apigw.ntruss.com/custom/v1/13003/d2bf285eb9d946e810d7f059f20746e61379e71acbc478ff8c871bdf88033dfb/infer"
    
 # 본인의 Secret Key로 치환
KEY = "eG1UY2tlcENmenpFWW5nemdZVk1Eb0ZNUnVtQnZIUnA="
    
headers = {
    "Content-Type": "application/json",
    "X-OCR-SECRET": KEY
}
    
data = {
    "version": "V1",
    "requestId": "sample_id", # 요청을 구분하기 위한 ID, 사용자가 정의
    "timestamp": 0, # 현재 시간값
    "images": [
        {
            "name": "sample_image",
            "format": "png",
            "data": img.decode('utf-8')
          # "templateIds": [400]  # 설정하지 않을 경우, 자동으로 템플릿을 찾음 
        }
    ]
}
data = json.dumps(data)
response = requests.post(URL, data=data, headers=headers)
res = json.loads(response.text)
print(json.dumps(res, sort_keys= True, indent=4, ensure_ascii = False))