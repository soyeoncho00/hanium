# -*- coding: utf-8 -*-
"""
Created on Wed May 11 00:37:23 2022

@author: Mgyu
"""

import win32com.client 
import pandas as pd

CPE_MARKET_KIND = {'KOSPI':1, 'KOSDAQ':2} # 코스피, 코스닥 지정
instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr") 

rows = list() 
for key, value in CPE_MARKET_KIND.items(): 
    codeList = instCpCodeMgr.GetStockListByMarket(value) # 주식 코드 목록 가져오기
    for code in codeList: 
        name = instCpCodeMgr.CodeToName(code) # 주식 종목 이름 가져오기
        sectionKind = instCpCodeMgr.GetStockSectionKind(code) #주식 종류 가저오기
        row = [code, name, key, sectionKind] 
        rows.append(row)

print('모든 종목을 불러왔습니다')

stockitems = pd.DataFrame(data= rows, columns=['code','name', 'section','sectionKind']) 

# sectionKind 값을 통해 ETF 구분하기
stockitems.loc[stockitems['sectionKind'] == 10, 'section'] = 'ETF' 

stockitems.to_csv('C:/Users/Mgyu/PL/financial/stockitems.csv', index=False)

print('파일을 저장하였습니다.')