import pythoncom
import win32com.client
import time
import statistics
import datetime

# Pycharm 에러에 대한 workaround - 'Unused import statement'
# noinspection PyUnresolvedReferences
import numpy as np
import pandas as pd
from pandas import DataFrame

class XASessionEvents:
    logInState = 0
    def OnLogin(self, code, msg):
        print("OnLogin method is called")
        print(str(code))
        print(str(msg))
        if str(code) == '0000':
            XASessionEvents.logInState = 1

    def OnLogout(self):
        print("OnLogout method is called")

    def OnDisconnect(self):
        print("OnDisconnect method is called")

class XAQueryEvents:
    queryState = 0
    def OnReceiveData(self, szTrCode):
        print("ReceiveData")
        XAQueryEvents.queryState = 1
    def OnReceiveMessage(self, systemError, messageCode, message):
        print("ReceiveMessage")


# T1305 호출을 위한 class
class XAQuery_t1305():

    def __init__(self):
        self.event  = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
        #self.event.parent = proxy(self)
        self.flag = False
        self.event.LoadFromResFile("C:\\eBEST\\xingAPI\\Res\\t1305.res")

    def Request(self,bNext=False):
        print("request called")
        self.event.Request(bNext)
        print("request called 2")
        self.flag = True
        print("request called 3")
        while self.flag:
            pythoncom.PumpWaitingMessages()
        print("request called 4")

    def SetFieldData(self, shcode):
        self.event.SetFieldData('t1305InBlock','shcode', 0, shcode) # 종목코드
        self.event.SetFieldData('t1305InBlock','dwmcode', 0, "1") # 일주월구분
        self.event.SetFieldData('t1305InBlock','cnt', 0, "1") # 날짜

    def GetFieldData(self,szBlockName,szFieldName,nOccur=-1):
        print("Get Field Data")
        if nOccur == -1:
            return self.event.GetFieldData(szBlockName,szFieldName)
        else:
            return self.event.GetFieldData(szBlockName,szFieldName,nOccur)

        self.dataReturn = [] # List
        self.idx = self.event.GetFieldData('t1305OutBlock','idx',0)

        nCount = self.event.GetBlockCount('t1305OutBlock1')
        for i in range(nCount):
            data = {} # Data Dictionary
            data['date'] = self.GetFieldData('t1305OutBlock1','date',i)
            data['open'] = str(self.GetFieldData('t1305OutBlock1','open',i))
            data['high'] = float(self.GetFieldData('t1305OutBlock1','high',i))
            data['low'] = float(self.GetFieldData('t1305OutBlock1','low',i))
            data['close'] = float(self.GetFieldData('t1305OutBlock1','close',i))
            data['value'] = float(self.GetFieldData('t1305OutBlock1','value',i))
            data['marketcap'] = float(self.GetFieldData('t1305OutBlock1','marketcap',i))

            self.dataReturn.append(data)

# ----------------------------------------------------------------------------
# login
# ----------------------------------------------------------------------------

id = input("아이디: ")
passwd = input("비밀번호: ")
cert_passwd = input("공인인증서: ")

# id = "아이디"
# passwd = "비밀번호"
# cert_passwd = "공인인증서"

instXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvents)
instXASession.ConnectServer("hts.ebestsec.co.kr", 20001)
instXASession.Login(id, passwd, cert_passwd, 0, 0)

while XASessionEvents.logInState == 0:
    pythoncom.PumpWaitingMessages()

# ----------------------------------------------------------------------------
# t1833 종목 가져오기
# ----------------------------------------------------------------------------
instXAQueryT1833 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
instXAQueryT1833.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t1833.res"

sFile = "C:\\eBEST\\xingAPI\\Res\\ConditionToApi_NEW.ADF"
instXAQueryT1833.RequestService("t1833", sFile)

while XAQueryEvents.queryState == 0:
    pythoncom.PumpWaitingMessages()

count = instXAQueryT1833.GetBlockCount("t1833OutBlock1")

dataList = []

for i in range(count):
    shcode = instXAQueryT1833.GetFieldData("t1833OutBlock1", "shcode", i)
    hname = instXAQueryT1833.GetFieldData("t1833OutBlock1", "hname", i)
    sign = instXAQueryT1833.GetFieldData("t1833OutBlock1", "sign", i)
    change = instXAQueryT1833.GetFieldData("t1833OutBlock1", "change", i)
    close = instXAQueryT1833.GetFieldData("t1833OutBlock1", "close", i)
    diff = instXAQueryT1833.GetFieldData("t1833OutBlock1", "diff", i)
    volume = instXAQueryT1833.GetFieldData("t1833OutBlock1", "volume", i)
    signcnt = instXAQueryT1833.GetFieldData("t1833OutBlock1", "signcnt", i)

    data = [shcode, hname, sign, change, close, diff, volume, signcnt]
    dataList.append(data)

stock = pd.DataFrame(dataList, columns=['종목코드', '종목명', '구분(5:하락, 2:상승)', '전일대비', '현재가', '등락율', '거래량', '연속봉수'])
#print("//종목 정보 출력")
print(stock)

# ----------------------------------------------------------------------------
# t1305 가격정보 추가: 기준일자(전일), 시가, 고가, 저가, 종가, 거래대금, 시가총액
# ----------------------------------------------------------------------------

# 1. TR코드 t1305로 종목코드의 기준일자(전일), 시가, 고가, 저가, 종가, 거래대금, 시가총액을 조회

#print("가격정보 연동 시작")

def GetData(shcode):
    XAQuery = XAQuery_t1305()
    print(shcode)
    XAQuery.SetFieldData(shcode)
    print("setfielddata")
    XAQuery.Request()
    print("request")
    return XAQuery.dataReturn

p_dataList = []

for index, row in stock.iterrows():

# XAQuery object는 종목코드만큼 생성해야 함
# 종목수만큼 object를 생성하고, eBest api 호출 --> 데이터를 포함한 object list 생성
# object list를 loop 돌면서 데이터를 추출하여, 데이터프레임을 생성함
    print("loop 시작")
    if __name__ == '__main__':
        data = GetData(row.ix[0])
        print(data)
        price = DataFrame(data, columns=['순서', '종목코드', '일자', '시가', '고가', '저가', '종가', '거래대금', '시가총액'])

    time.sleep(1)

#price = pd.DataFrame(p_dataList, columns=['순서', '종목코드', '일자', '시가', '고가', '저가', '종가', '거래대금', '시가총액'])
print(price)

'''
3. stock dataframe과 종목코드를 key로 merge
#selection = pd.merge(stock, price, on='shcode', how='inner')

# ----------------------------------------------------------------------------
# merge한 데이터프레임을 파일로 생성
# ----------------------------------------------------------------------------
#selection.to_csv('selection.csv', index=False)
'''