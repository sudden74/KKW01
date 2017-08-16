import pythoncom
import win32com.client
import xingLogin
import xingT1833
import xingT1305
import time
import sys

class XAQueryEvents:
    queryState = 0

    def OnReceiveData(self, szTrCode):
        print("ReceiveData" + str(szTrCode))
        XAQueryEvents.queryState = 1

    def OnReceiveMessage(self, systemError, messageCode, message):
        print("ReceiveMessage" + str(systemError) + str(messageCode) + str(message))

'''
def login():
    xingLogin.login()

def exeT1833():
    stock = xingT1833.getData()
    #print(stock)

    return stock
'''

if __name__ == '__main__':
### 로그인
    xingLogin.login()
    #login()

### 종목 선택
    #stock = exeT1833()
    stock = xingT1833.getData()
    #print(stock)

    #shcode = "043200"
    #print(shcode)

    for index, row in stock.iterrows():

        # 1안
        #1. 각 종목별로 XAQueryEvents 객체를 생성하고 eBest 데이터 연동(request 메쏘드 호출)
        #2. 종목(stocks) 리스트에 생성된 XAQueryEvents 객체를 append
        #3. stocks 리스트의 객체를 대상으로 loop, DataFrame을 생성

        # 2안
        #1. 각 종목별로 XAQueryEvents 객체를 생성
        #2. 종목(stocks) 리스트에 생성된 XAQueryEvents 객체를 append
        #3. stocks 리스트의 객체를 대상으로 loop, eBest 데이터를 연동하고(request 메쏘드 호출) DataFrame을 생성


        #instXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
        #stocks = [XAQueryEvents(row.ix[0])]
        #price = xingT1305.getData(row.ix[0])
        #price = xingT1305.getData(row.ix[0], instXAQuery)
        #print(price)
        #print(row.ix[0] + " reference count (after call): " + str(sys.getrefcount(XAQueryEvents)))

        time.sleep(1)
    # df.to_csv('kospi.csv')

