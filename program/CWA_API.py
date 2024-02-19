# -*- coding: utf-8 -*-
"""
Created on Mon Dec  4 11:04:16 2023

@author: User
"""

import pandas as pd
from datetime import datetime, timedelta
import os 

now = datetime.now()
yday = now - timedelta(days = 1)
today = now.strftime("%Y%m%d")
yesterday = yday.strftime("%Y%m%d")

path = rf"C:\Users\User\Desktop\ppt_自動化\yday_accumulate\{today}"


if os.path.exists(path):
    
    os.chdir(path)
    
else:
    
    os.mkdir(path)
    os.chdir(path)


url = "https://opendata.cwa.gov.tw/api/v1/rest/datastore/O-A0002-001?Authorization=rdec-key-123-45678-011121314"


df = pd.read_json(url)


data = {
        "基隆" : None,
        "三和" : None,
        "坪林" : None,
        "東澳" : None,
        "東澳嶺" : None,
        "金崙" : None,
        "桶後" : None,
        "巴陵" : None, 
        "池端" : None, 
        "南山、松茂" : None, 
        "德基、梨山" : None,
        "慈恩" : None,
        "廬山" : None,
        "廬山、合歡山" : None,
        "阿里山" : None,
        "甲仙" : None,
        "向陽" : None,
        "神木村" : None,
        "那瑪夏國中" : None,
        "加路蘭山" : None,
        "豐濱" : None
        }



id_dict = {
           "466940" : None, # 基隆
           "C0A930" : None, # 三和
           "C0A530" : None, # 坪林
           "C0U760" : None, # 東澳
           "C1U840" : None, # 東澳嶺
           "C0S790" : None, # 金崙
           "C0A570" : None, # 桶後
           "21C070" : None, # 巴陵
           "21U110" : None, # 池端
           "C0U720" : None, # 南山
           "81F860" : None, # 松茂
           "C1F9W0" : None, # 德基
           "C0F860" : None, # 梨山
           "C1T810" : None, # 慈恩
           "C0I010" : None, # 廬山
           "A0Z080" : None, # 合歡山
           "467530" : None, # 阿里山
           "C0V250" : None, # 甲仙
           "C0S750" : None, # 向陽
           "C0H9A0" : None, # 神木村
           "81V830" : None, # 那瑪夏國中
           "C0T9H0" : None, # 加路蘭山
           "C0T9I0" : None, # 豐濱
          }



def yday_rainfall(item):
    
    now = item["RainfallElement"]["Now"]["Precipitation"]
    past_two_day = item["RainfallElement"]["Past2days"]["Precipitation"]
    yday = past_two_day - now
    
    return yday



data_list = []


# 
for item in df.records.Station:
    
    dic = {}
    
    for id_stat in id_dict:
        
        if item["StationId"] == id_stat:
            
            id_dict[id_stat] = yday_rainfall(item)
            dic["id"] = id_stat
            dic["name"] = item["StationName"]
            dic["rainfall"] = id_dict[id_stat]            
            
            data_list.append(dic)


# 比較雨量大小
def compare(station, first_stat, second_stat):
    
    first_rainfall = 0
    second_rainfall = 0
    
    for i in data_list:
        
        if i["name"] == first_stat:
            
            first_rainfall = i["rainfall"]
        
        elif i["name"] == second_stat:
            
            second_rainfall = i["rainfall"]
            
    data[station] = max(first_rainfall, second_rainfall)
    
    
# 雨量資料寫入data
for i in data:
    
    if i == "南山、松茂":
        
        compare("南山、松茂", "南山", "松茂")
    
    elif i == "德基、梨山":
        
        compare("德基、梨山", "德基", "梨山")
        
    elif i == "廬山、合歡山":
        
        compare("廬山、合歡山", "廬山", "合歡山")
        
    else:
        
        for j in data_list:
            
            if j["name"] == i:
                
                data[i] = j["rainfall"]


# 轉置
def transpose(dic):
    
    # 轉成DataFrame
    df = pd.DataFrame(dic, index =  [f"昨日({yesterday})累積雨量"])
    
    # 轉置
    df_transposed = df.transpose()
    
    # 重新設置DataFrame索引，使其保持順序
    # df_transposed.reset_index(inplace=True)
    
    df_transposed.to_csv(rf'{yesterday}_累積雨量.csv', encoding="big5")
    
    
transpose(data)


