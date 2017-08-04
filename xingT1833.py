import pythoncom
import win32com.client

import pandas as pd

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
#stock.to_csv('T1833.csv')
