import requests
import json
login_url="http://bc.fjgdwl.net/login"
s=requests.session()
data={   
    "sessionId":"0",
    "username":"sjb",
    "verifyCode":"",
    "verifySMSVerifyCode":"",
    "authCode":"AA18B9EE69F51D49F0AE8045CDA69FB6"
    }
#r=requests.get(login_url)
headers = {'content-type': 'application/json'}  # payload请求方式
response=s.post(login_url,data=json.dumps(data),headers=headers)
print(response.text)