import requests
import json
import xlrd
import time
import sys

DEST_NODES='34888,35073,35039,34914,35063,35061,35045,35054,35058,35056,\
35044,34893,35020,35029,34910,34945,34940,34995,34997,34878,\
34986,34992,35003,34889,35009,34882,34931,34869,34962,34903,\
34981,34984,34961,34865,34970,35017,34920,35012,34993,35004,\
34971,35031,34982,34951,34975,34936,34948,34941,35001,34867,\
35000,34917,35005,34866,34928,35015,35010,35021,35025,34998,\
34968,34976,34955,35007,35016,34909,34956,34925,35018,35032,\
35028,34996,34911,34862,34924,35022,34988,34946,34880,34985,\
34989,34921,34934,34877,34977,34904,34943,34952,34875,34972,\
34959,34872,34923,34871,34895,35023,35037,34965,34873,34926,\
34912,34939,34974,34967,34884,34933,34958,34954,34960,34864,\
34887,35013,34944,34876,34860,34905,34994,34930,34879,34966,\
34885,34898,34938,34929,34907,34861,34883,35036,34896,34990,\
35014,34937,34973,34957,34991,34950,34987,34922,34891,34916,\
34979,34978,34919,34901,34868,34881,34963,34874,34899,34890,\
34953,34915,34894,35038,34918,34900,35033,34913,34999,35011,\
35035,34949,34983,35034,34863,34870,34964,34942,35006,34932,\
34892,34935,34927,35024,34902,34906,34908,34969,35026,34947,\
35002,35019,34980,34886,34897,35082,35080,35076,35078,35052,\
35064,35071,35083,35081,35086,35046,35079,35072,35077,35085,\
35060,35059,35066,35075,35043,35084,35047,35074,35042,35041,\
35069,35048,35053,35051,35049,35087,35055,35068,35040,35067,\
35062,35070,35065,35057,35050,35088'

login_url="http://bc.fjgdwl.net/login"
session=requests.session()
data={   
    "sessionId":"0",
    "username":"sjb",
    "verifyCode":"",
    "verifySMSVerifyCode":"",
    "authCode":"AA18B9EE69F51D49F0AE8045CDA69FB6"
    }
#r=requests.get(login_url)
headers = {'content-type': 'application/json'}  # payload请求方式
response=session.post(login_url,data=json.dumps(data),headers=headers)
print(response.text)
#coding:UTF-8


start_dt = "2018-02-18 00:00:00"
end_dt = "2018-02-24 23:59:59"
#转换成时间数组
start_timeArray = time.strptime(start_dt, "%Y-%m-%d %H:%M:%S")
end_timeArray = time.strptime(end_dt, "%Y-%m-%d %H:%M:%S")

#转换成时间戳
start_timestamp = str(int(time.mktime(start_timeArray)))
end_timestamp = str(int(time.mktime(end_timeArray)))

#print(start_timestamp)
#print(end_timestamp)

#print(sys.argv[1])
filename = sys.argv[1]
#print(filename)
workbook=xlrd.open_workbook(filename)
table=workbook.sheets()[0]
nrows = table.nrows

for row in range(nrows):
    #print(row)
    if(row == 0):
        continue
    source_node = str(int(table.cell(row,2).value))
    print(source_node)
    payload = {
        '_dc':str(int(time.time())),
        'graphId': '72',
        'testId': '98',
        'testType': '11',
        'start': '0',
        'limit': '1000',
        'reportDirection': '1',
        'exType': 'none',
        'exBeginTime': '0',
        'exEndTime': '0',
        'useSourceNode': 'true',
        'useDestNode': 'true',
        'beginTime': start_timestamp,
        'endTime': end_timestamp,
        'timeInterval': '604800',
        'sourceNodeIds': source_node,
        'destNodeIds':DEST_NODES,
        'indexNames': 'meanQuality',
        'indexTypes': 'AVG',
        'multiTests': '',
        'exportResult':'false',
        'meanQualityFormula':'0',
        'cloudTest':'',
        'conditions': '',
        'reportFlag':'true'
    }
    #print(payload)
    heads={}
    heads['User-Agent'] = 'Mozilla/5.0 ' \
                              '(Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' \
                              '(KHTML, like Gecko) Chrome/64.0.3282.167'
    r = session.get("http://bc.fjgdwl.net/getDataReportResult", params=payload,headers=headers)
    #print(r)
    retVal = (r.json()['results'].get(start_timestamp))
    #print(retVal)
    ret_sourceNode = retVal[0].get('sourceNodeId')
    qualityAVG = retVal[0].get('meanQuality_AVG')
   # print(ret_sourceNode)
    print('%.2f'% qualityAVG)
