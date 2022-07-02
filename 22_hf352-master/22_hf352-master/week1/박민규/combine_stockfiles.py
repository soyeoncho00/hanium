# -*- coding: utf-8 -*-
"""
Created on Wed May 11 03:11:04 2022

@author: Mgyu
"""

## 분할저장한 주가정보파일 합치기 
import pickle 
import pandas as pd 
n_file = 6 #파일 개수 
data = list() 
column_dailychart = ['code', 'section', 'date', 'open', 'high', 'low', 'close', 'vol', 'value', 
                     'n_stock', 'agg_price', 'foreign_rate','agency_buy', 'agency_netbuy'] 
for i in range(0, n_file): 
    with open('C:/Users/Mgyu/PL/financial/uncombined_files/dailychart{0}.txt'.format(i), 'rb') as f:
        tmp_data = pickle.load(f) 
        data = data + tmp_data 
dailychart = pd.DataFrame(data = data, columns= column_dailychart) 
dailychart[['value' ,'agg_price']] = (dailychart[['value' ,'agg_price']]/1000000).astype(int) 
dailychart = dailychart.sort_values(by=['code','date']) 
dailychart.to_csv('C:/Users/Mgyu/PL/financial/dailychart.csv', index=False)
