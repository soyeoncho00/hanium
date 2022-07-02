# -*- coding: utf-8 -*-
"""
Created on Wed May 11 00:44:30 2022

@author: Mgyu
"""

import win32com.client 
import pandas as pd 
import time 
import datetime 
column_dailychart = ['code', 'section', 'date', 'open', 'high', 'low', 'close', 'vol', 'value', 'n_stock', 
                     'agg_price', 'foreign_rate','agency_buy', 'agency_netbuy'] 

# 전에 만들었던 주식 목록 파일 불러오기
stockitems = pd.read_csv('C:/Users/Mgyu/PL/financial/stockitems.csv') 

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart") 
nCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")

row = list(range(len(column_dailychart))) 
rows = list() 
instStockChart.SetInputValue(1, ord('1'))
instStockChart.SetInputValue(2, (datetime.datetime.today() - datetime.timedelta(days=1)).strftime("%Y%m%d")) 
instStockChart.SetInputValue(3, '20180101') 
instStockChart.SetInputValue(5, (0, 2, 3, 4, 5, 8, 9, 12, 13, 17, 20, 21)) 
# 0: 날짜(ulong) 2: 시가(long or float) 3: 고가(long or float)
# 4: 저가(long or float) 5: 종가(long or float) 8: 거래량(ulong or ulonglong) 주) 정밀도 만원 단위
# 9: 거래대금(ulonglong) 12: 상장주식수(ulonglong) 13: 시가총액(ulonglong)
# 17: 외국인현보유비율(float) 20: 기관순매수(long) 21: 기관누적순매수(long)



instStockChart.SetInputValue(6, ord('D')) 
instStockChart.SetInputValue(9, ord('1'))

for idx, stockitem in stockitems.iterrows(): 
    remain_request_count = nCpCybos.GetLimitRemainCount(1) 
    print(stockitem['code'], stockitem['name'], '남은 요청 : ', remain_request_count) 
    if remain_request_count == 0: 
        print('남은 요청이 모두 소진되었습니다. 잠시 대기합니다.') 
        
        while True: 
            time.sleep(2) 
            remain_request_count = nCpCybos.GetLimitRemainCount(1) 
            if remain_request_count > 0: 
                print('작업을 재개합니다. (남은 요청 : {0})'.format(remain_request_count)) 
                break 
            print('대기 중...')
            
    instStockChart.SetInputValue(0, stockitem['code'])
    # BlockRequest
    instStockChart.BlockRequest() 
    # GetHeaderValue 
    numData = instStockChart.GetHeaderValue(3) 
    numField = instStockChart.GetHeaderValue(1) 
    # GetDataValue 
    for i in range(numData): 
        row[0] = stockitem['code'] 
        row[1] = stockitem['section'] # 코스피, 코스닥, ETF 여부 
        row[2] = instStockChart.GetDataValue(0, i) # 날짜 
        row[3] = instStockChart.GetDataValue(1, i) # 시가 
        row[4] = instStockChart.GetDataValue(2, i) # 고가 
        row[5] = instStockChart.GetDataValue(3, i) # 저가 
        row[6] = instStockChart.GetDataValue(4, i) # 종가 
        row[7] = instStockChart.GetDataValue(5, i) # 거래량 
        row[8] = instStockChart.GetDataValue(6, i) # 거래대금 
        row[9] = instStockChart.GetDataValue(7, i) # 상장주식수 
        row[10] = instStockChart.GetDataValue(8, i) # 시가총액 
        row[11] = instStockChart.GetDataValue(9, i) # 외국인 보율비율 
        row[12] = instStockChart.GetDataValue(10, i) # 기관순매수 
        row[13] = instStockChart.GetDataValue(11, i) # 기관누적순매수 
        rows.append(list(row)) 
        
        
print('데이터를 모두 불러왔습니다.') 

import pickle
unit = 500000 
for i in range(0, int(len(rows)/unit)): 
    with open('C:/Users/Mgyu/PL/financial/uncombined_files/dailychart{0}.txt'.format(i), 'wb') as f: 
        pickle.dump(rows[i*unit:(i+1)*unit], f) 
i+=1 
with open('C:/Users/Mgyu/PL/financial/uncombined_files/dailychart{0}.txt'.format(i), 'wb') as f: 
    pickle.dump(rows[(i)*unit:len(rows)], f) 
print('모든 데이터를 저장하였습니다.')



