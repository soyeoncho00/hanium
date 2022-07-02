# 2 ) 일봉 데이터 연속 조회 

# 키움증권의 KOA Studio를 참고해 일봉 데이터 요청에는 'opt10081'이라는 TR을 사용


import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import time

TR_REQ_TIME_INTERVAL = 0.2

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self._create_kiwoom_instance()
        self._set_signal_slots()

    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def _set_signal_slots(self): 
        self.OnEventConnect.connect(self._event_connect)
        self.OnReceiveTrData.connect(self._receive_tr_data) # 이벤트 발생 시 호출되게 하려면 시그널과 슬롯 연결

    def comm_connect(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

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

    def get_master_code_name(self, code):
        code_name = self.dynamicCall("GetMasterCodeName(QString)", code)
        return code_name

    def set_input_value(self, id, value): # SetInputValue 메서드를 통해 요청하는 TR에 필요한 데이터를 설정
        self.dynamicCall("SetInputValue(QString, QString)", id, value)

    def comm_rq_data(self, rqname, trcode, next, screen_no): # CommRqData 메서드를 호출해 TR을 키움증권 서버로 전송
        self.dynamicCall("CommRqData(QString, QString, int, QString", rqname, trcode, next, screen_no)
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_() # TR 요청 후 데이터 바로 반환되는 것이 아니라 이벤트 루프를 통해 키움증권이 이벤트를 줄 때까지 대기

    def _comm_get_data(self, code, real_type, field_name, index, item_name): # 키움증권 서버로부터 TR 처리에 대한 이벤트가 발생 시 실제로 데이터를 가져오기 위한 메서드
        ret = self.dynamicCall("CommGetData(QString, QString, QString, int, QString", code,
                               real_type, field_name, index, item_name)
        return ret.strip() # 해당 메서드 반환값이 양쪽에 공백이 있어 strip() 메서드로 제거

    def _get_repeat_cnt(self, trcode, rqname): # TR 요청 시 한꺼번에 너무 많은 데이터가 반환되므로, 몇 개의 데이터가 왔는지 알 수 있도록 함
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    def _receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4): # 이벤트가 발생했을 때 이를 처리하는 메서드
        if next == '2':
            self.remained_data = True
        else:
            self.remained_data = False

        if rqname == "opt10081_req":
            self._opt10081(rqname, trcode)

        try: # 더는 필요하지 않은 이벤트 루프를 종료
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    def _opt10081(self, rqname, trcode): 
        data_cnt = self._get_repeat_cnt(trcode, rqname) # 데이터 얻어오기 전 데이터의 개수를 해당 메서드를 통해 얻어옴

        for i in range(data_cnt):
            date = self._comm_get_data(trcode, "", rqname, i, "일자")
            open = self._comm_get_data(trcode, "", rqname, i, "시가")
            high = self._comm_get_data(trcode, "", rqname, i, "고가")
            low = self._comm_get_data(trcode, "", rqname, i, "저가")
            close = self._comm_get_data(trcode, "", rqname, i, "현재가")
            volume = self._comm_get_data(trcode, "", rqname, i, "거래량")
            print(date, open, high, low, close, volume)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.comm_connect()

    # opt10081 TR 요청
    kiwoom.set_input_value("종목코드", "039490")
    kiwoom.set_input_value("기준일자", "20170224")
    kiwoom.set_input_value("수정주가구분", 1)
    kiwoom.comm_rq_data("opt10081_req", "opt10081", 0, "0101")

    while kiwoom.remained_data == True:
        time.sleep(TR_REQ_TIME_INTERVAL) # 키움증권은 1초에 최대 5번의 TR 요청만 허용하므로 time 모듈의 sleep 함수를 통해 0.2초를 대기한 후 다음 TR 요청하도록 구현
        kiwoom.set_input_value("종목코드", "039490")
        kiwoom.set_input_value("기준일자", "20170224")
        kiwoom.set_input_value("수정주가구분", 1)
        kiwoom.comm_rq_data("opt10081_req", "opt10081", 2, "0101")