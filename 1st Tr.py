import pythoncom
import win32com.client
import time
import statistics
import datetime

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


# ----------------------------------------------------------------------------
# login
# ----------------------------------------------------------------------------
id = input( "아이디: ")
passwd = input( "비밀번호: ")
cert_passwd = input( "공인인증서: ")

#id = "아이디"
#passwd = "비밀번호"
#cert_passwd = "공인인증서"

instXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvents)
instXASession.ConnectServer("hts.ebestsec.co.kr", 20001)
instXASession.Login(id, passwd, cert_passwd, 0, 0)

while XASessionEvents.login_state == 0:
    pythoncom.PumpWaitingMessages()

# ----------------------------------------------------------------------------
# 종목파일을 읽어 데이터프레임을 생성
# ----------------------------------------------------------------------------
import pandas as pd
from pandas import DataFrame
import numpy as np

selection = pd.read_csv('selection.csv')


# ----------------------------------------------------------------------------
# 1차 매수 주문: CSPAT00600
# - 매수호가 = (저가 + 종가) /2
# - 수량 = 2,000,000 / 매수호가
# ----------------------------------------------------------------------------

inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
inXAQuery.LoadFromResFile("C:\\ETRADE\\xingAPI\\Res\\CSPAT00600.res")

inXAQuery.SetFieldData("CSPAT00600InBlock1", "AcntNo", 0, cbAccount)
inXAQuery.SetFieldData("CSPAT00600InBlock1", "InptPwd", 0, tbPwd)
inXAQuery.SetFieldData("CSPAT00600InBlock1", "IsuNo", 0, Range("L5").Value)
inXAQuery.SetFieldData("CSPAT00600InBlock1", "OrdQty", 0, Range("L9").Value)
inXAQuery.SetFieldData("CSPAT00600InBlock1", "OrdPrc", 0, Range("L10").Value)
inXAQuery.SetFieldData("CSPAT00600InBlock1", "BnsTpCode", 0, nOrderType)
inXAQuery.SetFieldData("CSPAT00600InBlock1", "OrdprcPtnCode", 0, Left(Range("L7").Value, 2))
inXAQuery.SetFieldData("CSPAT00600InBlock1", "MgntrnCode", 0, "000")
inXAQuery.SetFieldData("CSPAT00600InBlock1", "LoanDt", 0, "")
inXAQuery.SetFieldData("CSPAT00600InBlock1", "OrdCndiTpCode", 0, Left(Range("L8").Value, 1))
inXAQuery.Request(0)

while XAQueryEvents.queryState == 0:
    pythoncom.PumpWaitingMessages()

# 주문 번호 저장
name = inXAQuery.GetFieldData("CSPAT00600OutBlock2", "OrdNo", 0)
XAQueryEvents.queryState = 0