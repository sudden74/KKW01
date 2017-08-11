import pythoncom
import win32com.client
import pandas as pd
import sys

class XAQueryEvents:
    queryState = 0

    def OnReceiveData(self, szTrCode):
        print("ReceiveData")
        XAQueryEvents.queryState = 1

    def OnReceiveMessage(self, systemError, messageCode, message):
        print("ReceiveMessage")


# T1305 호출
def getData(shcode):
    
    print(shcode)
    '''
        # 변수 초기화 문제는 아님
            date = ''
            open = 0
            high = 0
            low = 0
            close = 0
            value = 0
            marketcap = 0
    '''
    instXAQueryT1305 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    instXAQueryT1305.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t1305.res"
    print(shcode + " reference count (before): " + str(sys.getrefcount(XAQueryEvents)))

    instXAQueryT1305.SetFieldData("t1305InBlock", "shcode", 0, shcode)  # 종목코드
    instXAQueryT1305.SetFieldData("t1305InBlock", "dwmcode", 0, "1")  # 일주월구분
    instXAQueryT1305.SetFieldData("t1305InBlock", "cnt", 0, "2")  # 날짜

    instXAQueryT1305.Request(0)

    while XAQueryEvents.queryState == 0:
        pythoncom.PumpWaitingMessages()

    count = instXAQueryT1305.GetBlockCount("T1305OutBlock1")

    dataList = []

    for i in range(count):
        if i == 1:
            date = instXAQueryT1305.GetFieldData("t1305OutBlock1", "date", 0)  # 일자
            open = instXAQueryT1305.GetFieldData("t1305OutBlock1", "open", 0)  # 시가
            high = instXAQueryT1305.GetFieldData("t1305OutBlock1", "high", 0)  # 고가
            low = instXAQueryT1305.GetFieldData("t1305OutBlock1", "low", 0)  # 저가
            close = instXAQueryT1305.GetFieldData("t1305OutBlock1", "close", 0)  # 종가
            value = instXAQueryT1305.GetFieldData("t1305OutBlock1", "value", 0)  # 거래대금
            marketcap = instXAQueryT1305.GetFieldData("t1305OutBlock1", "marketcap", 0)  # 시가총액

            print(shcode, date, open, high, low, close, value, marketcap)
            data = [shcode, date, open, high, low, close, value, marketcap]
            #print(data)
            dataList.append(data)


    print(dataList)
    price = pd.DataFrame(dataList, columns=['종목코드', '일자', '시가', '고가', '저가', '종가', '거래대금', '시가총액'])
    print(price)
    #print("//종목 정보 출력")
    #print(stock)
    #stock.to_csv('T1305.csv')
    return price

    del instXAQueryT1305

    print(shcode + " reference count (after): " + str(sys.getrefcount(XAQueryEvents)))
