import xingLogin
import xingT1833
import xingT1305
import time

def login():
    xingLogin.login()

def exeT1833():
    stock = xingT1833.getData()
    #print(stock)

    return stock

if __name__ == '__main__':
### 로그인
    login()

### 종목 선택
    stock = exeT1833()
    #stock = xingT1833.getData()
    print(stock)

    #shcode = "043200"
    #print(shcode)

    for index, row in stock.iterrows():

        price = xingT1305.getData(row.ix[0])
        #print(price)

        time.sleep(1)
    # df.to_csv('kospi.csv')

