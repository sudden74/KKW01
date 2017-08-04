import pythoncom
import win32com.client


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
        self.event.Request(bNext)
        self.flag = True
        while self.flag:
            pythoncom.PumpWaitingMessages()

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