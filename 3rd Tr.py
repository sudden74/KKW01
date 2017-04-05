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
    def OnReceiveMessage(self, systemError, messageCode, message):
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
