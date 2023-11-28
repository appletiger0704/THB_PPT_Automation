# -*- coding: utf-8 -*-
"""
Created on Wed Nov 22 15:08:46 2023

@author: User
"""

import pandas as pd
import os
from datetime import datetime, timedelta


now = datetime.now()
ydate = now - timedelta(days = 1)


today = now.strftime("%Y%m%d")
yday = ydate.strftime("%Y%m%d")


path = rf"C:\Users\User\Desktop\ppt_自動化\yday_accumulate\{today}"


if os.path.exists(path):
    
    os.chdir(path)
    
else:
    
    os.mkdir(path)
    os.chdir(path)


CWA_URL = "https://opendata.cwa.gov.tw/api/v1/rest/datastore/O-A0002-001?Authorization=CWA-0B81E3AE-B6D1-4225-9B06-A4A6647E3998&format=JSON&StationId=466940,C0A930,C0A530,C0U760,C1U840,C0S790,C0A570,C0U720,C1F9W0,C0F860,C1T810,C0I010,C0H9C0,467530,C0V250,C0S750,C0H9A0,C0T9H0,C0T9I0&GeoInfo=TownCode"


WRA_URL = {
    "巴陵" : "https://fhy.wra.gov.tw/Api/v2/Rainfall/Info/Last24Hours/StationNo/21C070/H24",  
    "池端" : "https://fhy.wra.gov.tw/Api/v2/Rainfall/Info/Last24Hours/StationNo/21U110/H24", 
    "松茂" : "https://fhy.wra.gov.tw/Api/v2/Rainfall/Info/Last24Hours/StationNo/81F860/H24", 
    "那瑪夏國中" :"https://fhy.wra.gov.tw/Api/v2/Rainfall/Info/Last24Hours/StationNo/81V830/H24"   
    }



data = {
        "基隆" : None,
        "三和" : None,
        "坪林" : None,
        "東澳" : None,
        "東澳嶺" : None,
        "金崙" : None,
        "桶後" : None,
        "巴陵" : None, # 水利署測站
        "池端" : None, # 水利署測站
        "南山、松茂" : None, # 松茂是水利署測站
        "德基、梨山" : None,
        "慈恩" : None,
        "廬山" : None,
        "廬山、合歡山" : None,
        "阿里山" : None,
        "甲仙" : None,
        "向陽" : None,
        "神木村" : None,
        "那瑪夏國中" : None, # 水利署測站
        "加路蘭山" : None,
        "豐濱" : None
        }



# wra_data = {
#             "巴陵" : None,
#             "池端" : None,
#             "松茂" : None,
#             "那瑪夏國中" : None
#             }


def write_txt(string):
    
    with open(f"{today}_log.txt", "a") as file:
        
        file.write(f"{now.hour}時{now.minute}分 \n" + string + "\n")


# 轉置矩陣DataFrame，並存成csv檔案
def transpose(dic):
    
    # 轉成DataFrame
    df = pd.DataFrame(dic, index =  [f"昨日({yday})累積雨量"])
    
    # 轉置
    df_transposed = df.transpose()
    
    # 重新設置DataFrame索引，使其保持順序
    # df_transposed.reset_index(inplace=True)
    
    df_transposed.to_csv(rf'{yday}_累積雨量.csv', encoding="big5")
    


# wra json data 
def time_count(raw_data):
    
    count = -1
        
    for j in raw_data["Data"][0]["Timeseries"]:
        
        count = count + 1
        
        if j == "0:0":
            
            return count


# 紀錄水利署雨量站雨量值到 wra_data
def wra_rainfall(name, count, raw_data):
    
    rainfall_array = raw_data["Data"][0]["Rainfall"]
    
    if name == "松茂":
        
        data["南山、松茂"] = rainfall_array[count]
    
    else:
        
        data[name] = rainfall_array[count]


# 執行程式
def wra_exe():
    
    for key, url in WRA_URL.items():
        
        try:
            
            raw_data = pd.read_json(url)
        
        except Exception as e:
            
            write_txt(e)
        
        wra_rainfall(key, time_count(raw_data), raw_data)
        
        print(key + "(水利署雨量站) _資料時間" +  raw_data["Data"][0]["Time"])
        
        
wra_exe()


# 嘗試讀取API
try:
    
    raw_data = pd.read_json(CWA_URL)

# 錯誤處理
except Exception as e:
    
    write_txt(e)


# 擷取CWA雨量站資料
for i in raw_data.records.Station:
    
    # 計算昨日累積雨量
    past2days = i["RainfallElement"]["Past2days"]["Precipitation"]
    now = i["RainfallElement"]["Now"]["Precipitation"]
    yday_rainfall = past2days - now
    
    # 紀錄當前雨量站名稱
    station_name = i["StationName"]
    
    
    
    # 德基、梨山 雨量站比較，紀錄較大值
    Deji_Lishan_Rainfall = 0
    
    if station_name == "德基" or station_name == "梨山" :
        
        if yday_rainfall >= Deji_Lishan_Rainfall:
            
            Deji_Lishan_Rainfall = yday_rainfall
            
        data["德基、梨山"] = Deji_Lishan_Rainfall
        
        
        
    # 廬山、合歡山 雨量站比較，紀錄較大值
    Lushan_Hehuan_Rainfall = 0
    
    if station_name == "廬山" or station_name == "合歡山" :
        
        if yday_rainfall >= Lushan_Hehuan_Rainfall:
            
            Lushan_Hehuan_Rainfall = yday_rainfall
            
        data["廬山、合歡山"] = Lushan_Hehuan_Rainfall
        
    
    
    # 南山、松茂(水利署雨量站) 雨量值比較
    if station_name == "南山":
        
        if yday_rainfall >= data["南山、松茂"]:
            
            data["南山、松茂"] = yday_rainfall
            
            

    # 將昨日累積雨量記錄到data
    for k in data.keys():
                
        if k == station_name:
            
            data[k] = yday_rainfall
    
    # API資料時間
    data_time = i["ObsTime"]["DateTime"]
    
    print(f"{station_name} _ 資料時間：{data_time}")
        


transpose(data)