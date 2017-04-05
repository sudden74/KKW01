import pythoncom
import win32com.client
import time
import statistics
import datetime

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
        XAQueryEventHandlerT1833.query_state = 1



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
id = input( "아이디: ")
passwd = input( "비밀번호: ")
cert_passwd = input( "공인인증서: ")

#id = "아이디"
#passwd = "비밀번호"
#cert_passwd = "공인인증서"

instXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEventHandler)
instXASession.ConnectServer("hts.ebestsec.co.kr", 20001)
instXASession.Login(id, passwd, cert_passwd, 0, 0)

while XASessionEventHandler.login_state == 0:
    pythoncom.PumpWaitingMessages()

# ----------------------------------------------------------------------------
# t1305 가격정보 추가: 기준일자(전일), 시가, 고가, 저가, 종가, 거래대금, 시가총액
# ----------------------------------------------------------------------------

    #1. TR코드 t1305로 종목코드의 기준일자(전일), 시가, 고가, 저가, 종가, 거래대금, 시가총액을 조회

p_dataList = []

import pandas as pd
from pandas import DataFrame
import numpy as np

raw_data = {'종목코드':['075130', '095300', '101140']}
stock = DataFrame(raw_data)

'''
instXAQueryT1305 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT1305)
instXAQueryT1305.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t1305.res"
'''


#for index, row in stock.iterrows():

time.sleep(1)

#    print(row['종목코드'])

instXAQueryT1305 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT1305)
instXAQueryT1305.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t1305.res"

    shcode = '075130'

#    print(shcode)

instXAQueryT1305.SetFieldData("t1305InBlock", "shcode", 0, shcode) #종목코드
instXAQueryT1305.SetFieldData("t1305InBlock", "dwmcode", 0, "1") #일주월구분
instXAQueryT1305.SetFieldData("t1305InBlock", "cnt", 0, "2") #날짜

    #t1305 요청
instXAQueryT1305.Request(False)

while XAQueryEventHandlerT1305.query_state == 0:
    pythoncom.PumpWaitingMessages()


date = instXAQueryT1305.GetFieldData("t1305OutBlock1", "date", 1)        #일자
open = instXAQueryT1305.GetFieldData("t1305OutBlock1", "open", 1)        #시가
high = instXAQueryT1305.GetFieldData("t1305OutBlock1", "high", 1)        #고가
low = instXAQueryT1305.GetFieldData("t1305OutBlock1", "low", 1)        #저가
close = instXAQueryT1305.GetFieldData("t1305OutBlock1", "close", 1)        #종가
value = instXAQueryT1305.GetFieldData("t1305OutBlock1", "value", 1)        #거래대금
marketcap = instXAQueryT1305.GetFieldData("t1305OutBlock1", "marketcap", 1)        #시가총액

print(shcode, date, open, high, low, close, value, marketcap)

    #2. price dataframe 생성
p_data = [shcode, date, open, high, low, close, value, marketcap]
print(p_data)


'''
    p_dataList.append(p_data)



price = pd.DataFrame(p_dataList, columns=['종목코드', '일자', '시가', '고가', '저가', '종가', '거래대금', '시가총액'])
print(price)

'''