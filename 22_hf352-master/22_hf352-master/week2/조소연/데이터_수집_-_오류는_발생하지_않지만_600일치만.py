import sys
from PyQt5.QtWidgets import *
import Kiwoom
import time
import csv
from pandas import DataFrame
import os

MARKET_KOSPI   = 0
MARKET_KOSDAQ  = 10

class PyMon:
    def __init__(self):
        self.kiwoom = Kiwoom.Kiwoom()
        self.kiwoom.comm_connect()
        self.get_code_list()

    def get_code_list(self):
        self.kospi_codes = self.kiwoom.get_code_list_by_market(MARKET_KOSPI)
        self.kosdaq_codes = self.kiwoom.get_code_list_by_market(MARKET_KOSDAQ)


    def get_ohlcv(self, code, start):
        self.kiwoom.ohlcv = {'date': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': []}

        self.kiwoom.set_input_value("종목코드", code)
        self.kiwoom.set_input_value("기준일자", start)
        self.kiwoom.set_input_value("수정주가구분", 1)


        self.kiwoom.comm_rq_data("opt10081_req", "opt10081", 0, "0101")
        # time.sleep(0.2)
        time.sleep(3.6)


        df = DataFrame(self.kiwoom.ohlcv, columns=['open', 'high', 'low', 'close', 'volume'],
                       index=self.kiwoom.ohlcv['date'])
        return df

    def run(self):

        # for kospi_code in self.kospi_codes:
        #     itemName = self.kiwoom.get_master_code_name(kospi_code)
        #     df = self.get_ohlcv(kospi_code, "20220516")
        #     df.to_csv('./Data/kospi/' + itemName + '.csv')


        for kosdaq_code in self.kosdaq_codes:
            itemName = self.kiwoom.get_master_code_name(kosdaq_code)
            df = self.get_ohlcv(kosdaq_code, "20220516")
            df.to_csv('./Data/kosdaq/' + itemName + '.csv')


        # df = self.get_ohlcv("039490", "20220515")
        # print(df)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    pymon = PyMon()
    pymon.run()