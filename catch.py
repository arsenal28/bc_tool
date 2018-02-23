from urllib import urlopen
import requests
from bs4 import BeautifulSoup
import sys
import xlwt
import xlrd

s = requests.session()
user='sjb'
pwd='AA18B9EE69F51D49F0AE8045CDA69FB6'
para = {'username':user,'verifyCode':'','verifySMSVerifyCode':'','authCode':pwd}
url = 'http://192.168.63.43/login'
r =s.post(url, data=para)
print(r.text)
#soup = BeautifulSoup(r.text,'html.parser')
