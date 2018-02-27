import requests
import json
import xlrd
import time
import sys
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
end_dt = "2018-02-25 00:00:00"
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
    graph_id = str(int(table.cell(row,4).value))
    test_id = str(int(table.cell(row,5).value))
    node_id = str(int(table.cell(row,6).value))
    dest_nodes = str(table.cell(row,7))
    #要填充报告的行列
    cell_row = int(table.cell(row,2).value)
    cell_col = int(table.cell(row,3).value)

    payload = {
        '_dc':str(int(time.time())),
        'graphId': graph_id,
        'testId': test_id,
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
        'sourceNodeIds': node_id,
        'destNodeIds':dest_nodes,
        'indexNames': 'meanQuality',
        'indexTypes': 'AVG',
        'multiTests': '',
        'exportResult':'false',
        'meanQualityFormula':'0',
        'cloudTest':'',
        'conditions': '',
        'reportFlag':'true'
    }
    headers={}
    headers['User-Agent'] = 'Mozilla/5.0 ' \
                              '(Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' \
                              '(KHTML, like Gecko) Chrome/64.0.3282.167'
    r = session.get("http://bc.fjgdwl.net/getDataReportResult", params=payload,headers=headers)
    retVal = (r.json()['results'].get(start_timestamp))
    ret_sourceNode = retVal[0].get('sourceNodeId')
    qualityAVG = retVal[0].get('meanQuality_AVG')
    print('%.2f'% qualityAVG)
