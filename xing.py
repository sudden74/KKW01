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

        #instXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
        price = xingT1305.getData(row.ix[0])
        #price = xingT1305.getData(row.ix[0], instXAQuery)
        print(price)
        #print(row.ix[0] + " reference count (after call): " + str(sys.getrefcount(XAQueryEvents)))

        time.sleep(1)
    # df.to_csv('kospi.csv')

