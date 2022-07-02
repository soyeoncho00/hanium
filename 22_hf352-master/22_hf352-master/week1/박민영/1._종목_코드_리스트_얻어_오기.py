# 1 ) 종목 코드 리스트 얻어 오기

# GetCodeListByMarket 메서드 : 각 시장에 속하는 종목의 종목 코드 리스트를 얻을 수 있음
# 0 : 장내 / 3 : ELW / 4 : 뮤추얼펀드 / 5 : 신주인수권 / 6 : 리츠 /
# 8 : ETF / 9 : 하이일드펀드 / 10 : 코스닥 / 30 : 제3시장

# Kiwoom 클래스 : 파이썬에서 키움 OpenAPI+를 쉽게 사용할 수 있게 해주는 역할을 하는 클래스

import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *

class Kiwoom(QAxWidget): # Kiwoon 클래스는 QAxWidget 클래스를 상속받음으로써 Kiwoom 클래스의 인스턴스가 QAxWidget 클래스에서 제공하는 메서드를 호출할 수 있게 함
    def __init__(self):
        super().__init__()
        self._create_kiwoom_instance()
        self._set_signal_slots()

    def _create_kiwoom_instance(self): # 키움증권 OpenAPI+ 사용 위해 COM 오브젝트 생성
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def _set_signal_slots(self): # 키움증권 서버로부터 발생한 이벤트와 이를 처리할 메서드를 연결
        self.OnEventConnect.connect(self._event_connect)

    def comm_connect(self): # 키움증권의 OpenAPI+에 로그인
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_() # 이벤트 루프 생성, 이벤트가 발생할 때까지 프로그램이 종료되지 않게 함

    def _event_connect(self, err_code):
        if err_code == 0:
            print("connected")
        else:
            print("disconnected")

        self.login_event_loop.exit()

    def get_code_list_by_market(self, market):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market)
        code_list = code_list.split(';')
        return code_list[:-1]

if __name__ == "__main__":
    app = QApplication(sys.argv) # Kiwoom 클래스는 QAxWidget 클래스를 상속받았기 때문에 QApplication 클래스의 인스턴스를 먼저 생성
    kiwoom = Kiwoom() # 추후 Kiwoom 클래스에 대한 인스턴스 생성 가능
    kiwoom.comm_connect() # Kiwoom 객체 생성 후 comm_connect() 메서드 호출해 로그인 수행
    code_list = kiwoom.get_code_list_by_market('10') # 코스닥 시장은 10에 해당
    for code in code_list:
        print(code, end=" ")