# -*- coding: UTF-8 -*-
import requests
import json
import xlrd
import xlwt
import time,datetime
import sys
import shutil
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
if (len(sys.argv) != 3) and (len(sys.argv) != 1):
    print ('输入参数错误！  输入格式如下 python bc_tool.py 20180218 20180224.')
    sys.exit()
if (len(sys.argv) == 1):
    end_date= time.strftime("%Y%m%d")
    print(end_date)
    start_date = (datetime.datetime.today() - datetime.timedelta(days=6)).strftime("%Y%m%d")
    print(start_date)
if (len(sys.argv) == 3):
    start_date = sys.argv[1]
    end_date = sys.argv[2]
start_dt = str(start_date)+" "+"00:00:00"
end_dt = str(end_date)+" "+"23:59:59"
#转换成时间数组
start_timeArray = time.strptime(start_dt, "%Y%m%d %H:%M:%S")
end_timeArray = time.strptime(end_dt, "%Y%m%d %H:%M:%S")
#转换成时间戳
start_timestamp = str(int(time.mktime(start_timeArray)))
end_timestamp = str(int(time.mktime(end_timeArray)))

report_period = str(start_timeArray[1])+'月'+str(start_timeArray[2])+'日'+'-'+str(end_timeArray[1])+'月'+str(end_timeArray[2])+'日'

#print(start_timestamp)
#print(end_timestamp)

#print(sys.argv[1])
filename = 'bc.xlsx'

#print(filename)
workbook=xlrd.open_workbook(filename)
table=workbook.sheets()[0]
nrows = table.nrows
bc_book=xlwt.Workbook(encoding='utf-8')
bc_sheet = bc_book.add_sheet('bc_sheet', cell_overwrite_ok=True)
##############################################################################################
font = xlwt.Font() # Create the Font
font.name = u'宋体'
font.bold = True
font.height = 200
borders = xlwt.Borders()
borders.left = xlwt.Borders.THIN
borders.right = xlwt.Borders.THIN
borders.top = xlwt.Borders.THIN
borders.bottom = xlwt.Borders.THIN
alignment = xlwt.Alignment() # Create Alignment
alignment.horz = xlwt.Alignment.HORZ_CENTER
alignment.vert = xlwt.Alignment.VERT_CENTER
style = xlwt.XFStyle()
style.alignment = alignment # Add Alignment to Style
style.font = font
style.borders = borders
bc_sheet.write_merge(1, 2, 0, 0, '日期',style)
bc_sheet.write_merge(1, 1, 2, 3, '电信出口',style)
bc_sheet.write_merge(1, 1, 4, 5, '省联通出口',style)
bc_sheet.write_merge(1, 1, 6, 7, '华数出口',style)
bc_sheet.write_merge(1, 1, 8, 9, '东方网信出口',style)
bc_sheet.write_merge(1, 1, 10, 11, '网宿出口',style)
bc_sheet.write_merge(1, 1, 12, 13, '福州分公司'+'\n'+'企舜混合出口',style)
bc_sheet.write_merge(1, 1, 14, 15, '福州分公司'+'\n'+'华数混合出口',style)
bc_sheet.write_merge(1, 2, 1, 1, '分公司',style)
bc_sheet.write_merge(3, 12, 0, 0, report_period,style)
x=2
while x<=14:
    bc_sheet.write(2 ,x,'平均',style)
    bc_sheet.write(2 ,x+1,'忙时',style)
    x+=2
bc_sheet.write(3 ,1,'集团',style)
bc_sheet.write(4 ,1,'福州',style)
bc_sheet.write(5 ,1,'厦门',style)
bc_sheet.write(6 ,1,'宁德',style)
bc_sheet.write(7 ,1,'莆田',style)
bc_sheet.write(8 ,1,'泉州',style)
bc_sheet.write(9 ,1,'漳州',style)
bc_sheet.write(10,1,'龙岩',style)
bc_sheet.write(11,1,'三明',style)
bc_sheet.write(12,1,'南平',style)

first_col=bc_sheet.col(0)
first_col.width=256*20
################################################################################################
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
    qualityAVG = round(retVal[0].get('meanQuality_AVG'),1)
    #print('%.2f'% qualityAVG)
    bc_sheet.write(cell_row,cell_col,qualityAVG)

    payload_busy = {
        '_dc':str(int(time.time())),
        'graphId': graph_id,
        'testId': test_id,
        'testType': '11',
        'start': '0',
        'limit': '1000',
        'reportDirection': '1',
        'exType': 'hour',
        'exBeginTime': '18',
        'exEndTime': '22',
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
    r = session.get("http://bc.fjgdwl.net/getDataReportResult", params=payload_busy,headers=headers)

    retVal = (r.json()['results'].get(start_timestamp))
    ret_sourceNode = retVal[0].get('sourceNodeId')
    qualityAVG = round(retVal[0].get('meanQuality_AVG'),1)
    #print('%.2f'% qualityAVG)
    bc_sheet.write(cell_row,cell_col+1,qualityAVG)
report_name = 'bc_report_'+start_date+'-'+end_date+'.xls'
bc_book.save(report_name)
#report_name_with_path = '/var/www/html/zhoubao/'+report_name
#shutil.copyfile(report_name,report_name_with_path)