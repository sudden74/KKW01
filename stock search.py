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


class XASessionEventHandler:
    login_state = 0

    def OnLogin(self, code, msg):
        if code == "0000":
            print("로그인 성공")
            XASessionEventHandler.login_state = 1
        else:
            print("로그인 실패")


class XAQueryEventHandlerT1833:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT1833.query_state = 1


class XAQueryEventHandlerT1305:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandlerT1305.query_state = 1


'''
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
    def OnReceiveMessage(self, systemError, mesageCode, message):
        print("ReceiveMessage")
'''

# ----------------------------------------------------------------------------
# login
# ----------------------------------------------------------------------------
id = input("아이디: ")
passwd = input("비밀번호: ")
cert_passwd = input("공인인증서: ")

# id = "아이디"
# passwd = "비밀번호"
# cert_passwd = "공인인증서"

instXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEventHandler)
instXASession.ConnectServer("hts.ebestsec.co.kr", 20001)
instXASession.Login(id, passwd, cert_passwd, 0, 0)

while XASessionEventHandler.login_state == 0:
    pythoncom.PumpWaitingMessages()

# ----------------------------------------------------------------------------
# t1833 종목 가져오기
# ----------------------------------------------------------------------------
instXAQueryT1833 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT1833)
instXAQueryT1833.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t1833.res"

sFile = "C:\\eBEST\\xingAPI\\Res\\ConditionToApi_NEW.ADF"
instXAQueryT1833.RequestService("t1833", sFile)

while XAQueryEventHandlerT1833.query_state == 0:
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
print("//종목 정보 출력")
print(stock)

# ----------------------------------------------------------------------------
# t1305 가격정보 추가: 기준일자(전일), 시가, 고가, 저가, 종가, 거래대금, 시가총액
# ----------------------------------------------------------------------------

# 1. TR코드 t1305로 종목코드의 기준일자(전일), 시가, 고가, 저가, 종가, 거래대금, 시가총액을 조회

j=0
p_dataList = []

print("가격정보 연동 시작")

for index, row in stock.iterrows():

# XAQuery object는 종목코드만큼 생성해야 함
# 종목수만큼 object를 생성하고, eBest api 호출 --> 데이터를 포함한 object list 생성
# object list를 loop 돌면서 데이터를 추출하여, 데이터프레임을 생성함

    j = j + 1
    instXAQueryT1305 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT1305)
    instXAQueryT1305.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t1305.res"

    print("start of loop")

    print(j, row.ix[0])

# SetFieldData가 변경되지 않음
# instXAQueryT1305를 계속해서 terminate하고 새로 생성해주어야 하나?
    instXAQueryT1305.SetFieldData("t1305InBlock", "shcode", 0, row.ix[0])  # 종목코드
    instXAQueryT1305.SetFieldData("t1305InBlock", "dwmcode", 0, "1")  # 일주월구분
    instXAQueryT1305.SetFieldData("t1305InBlock", "cnt", 0, "1")  # 날짜

    # t1305 요청
    instXAQueryT1305.Request(False)

    while XAQueryEventHandlerT1305.query_state == 0:
        pythoncom.PumpWaitingMessages()

    date = instXAQueryT1305.GetFieldData("t1305OutBlock1", "date", 0)  # 일자
    open = instXAQueryT1305.GetFieldData("t1305OutBlock1", "open", 0)  # 시가
    high = instXAQueryT1305.GetFieldData("t1305OutBlock1", "high", 0)  # 고가
    low = instXAQueryT1305.GetFieldData("t1305OutBlock1", "low", 0)  # 저가
    close = instXAQueryT1305.GetFieldData("t1305OutBlock1", "close", 0)  # 종가
    value = instXAQueryT1305.GetFieldData("t1305OutBlock1", "value", 0)  # 거래대금
    marketcap = instXAQueryT1305.GetFieldData("t1305OutBlock1", "marketcap", 0)  # 시가총액

    print("//연동 값 출력")
    print(i, shcode, date, open, high, low, close, value, marketcap)

    # 2. price dataframe 생성
    p_data = [i, shcode, date, open, high, low, close, value, marketcap]
    print("//가격 리스트 출력")
    print(p_data)

    p_dataList.append(p_data)
    print("//가격 데이터프레임 출력")
    print(p_dataList)

    print("end of loop")

    time.sleep(1)

price = pd.DataFrame(p_dataList, columns=['순서', '종목코드', '일자', '시가', '고가', '저가', '종가', '거래대금', '시가총액'])
print(price)
'''
# 3. stock dataframe과 종목코드를 key로 merge
selection = pd.merge(stock, price, on='shcode', how='inner')

# ----------------------------------------------------------------------------
# merge한 데이터프레임을 파일로 생성
# ----------------------------------------------------------------------------
selection.to_csv('selection.csv', index=False)
'''