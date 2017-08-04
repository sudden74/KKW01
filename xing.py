import xingLogin
import xingT1833
import xingT1305
import time

if __name__ == '__main__':
### 로그인
    xingLogin.login()

### 종목 선택
    stock = xingT1833.getData()
    print(stock)

    #shcode = "043200"
    #print(shcode)

    for index, row in stock.iterrows():

        price = xingT1305.getData(row.ix[0])
        #print(price)

        time.sleep(1)
    # df.to_csv('kospi.csv')


'''
time.sleep(1)

data2=GetData('2', 'a')
print(data)
df2=DataFrame(data2) #columns=['Date', 'Open', 'High', 'Low', 'Close'])
df.to_csv('kosdaq.csv')
'''
