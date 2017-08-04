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
# t1305 가격정보 추가: 기준일자(전일), 시가, 고가, 저가, 종가, 거래대금, 시가총액
# ----------------------------------------------------------------------------

# 1. TR코드 t1305로 종목코드의 기준일자(전일), 시가, 고가, 저가, 종가, 거래대금, 시가총액을 조회

code = "078940"

instXAQueryT1305 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandlerT1305)
instXAQueryT1305.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t1305.res"

#instXAQueryT1305.SetFieldData("t1305InBlock", "shcode", 0, code)  # 종목코드
instXAQueryT1305.SetFieldData("t1305InBlock", "shcode", 0, code)  # 종목코드
instXAQueryT1305.SetFieldData("t1305InBlock", "dwmcode", 0, "1")  # 일주월구분
instXAQueryT1305.SetFieldData("t1305InBlock", "cnt", 0, "2")  # 날짜

# t1305 요청
instXAQueryT1305.Request(0)

while XAQueryEventHandlerT1305.query_state == 0:
    pythoncom.PumpWaitingMessages()

#count = instXAQueryT1305.GetBlockCount("t1305OutBlock")
#for i in range(count):
date = instXAQueryT1305.GetFieldData("t1305OutBlock1", "date", 1)  # 일자
open = instXAQueryT1305.GetFieldData("t1305OutBlock1", "open", 1)  # 시가
high = instXAQueryT1305.GetFieldData("t1305OutBlock1", "high", 1)  # 고가
low = instXAQueryT1305.GetFieldData("t1305OutBlock1", "low", 1)  # 저가
close = instXAQueryT1305.GetFieldData("t1305OutBlock1", "close", 1)  # 종가
value = instXAQueryT1305.GetFieldData("t1305OutBlock1", "value", 1)  # 거래대금
marketcap = instXAQueryT1305.GetFieldData("t1305OutBlock1", "marketcap", 1)  # 시가총액

print(code, date, open, high, low, close, value, marketcap)
