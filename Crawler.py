# -*- coding: utf-8 -*-
"""
Created on Mon Nov  6 17:39:45 2023

@author: User
"""

import os
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from io import BytesIO
from PIL import Image
from datetime import datetime, timedelta
import re


now = datetime.now()
one_day = timedelta(days = 1)
one_hour = timedelta(hours = 1)
one_hours_ago = now - one_hour
# yesterday = now - one_day
# next_day = now + one_day
# day_after_tomorrow = now + 2*one_day

# 圖片儲存位置
path = rf"C:\Users\User\Desktop\ppt_自動化\{now.year}{now.month}{now.day}"

if not os.path.exists(path):
    os.mkdir(path)

os.chdir(path)

# 會使用到的網址
URL ={
      "SWM" : "https://www.cwa.gov.tw/V8/C/W/analysis.html",
      "Radar" : "https://www.cwa.gov.tw/V8/C/W/OBS_Radar.html",
      "StreamLine" : "https://tropic.ssec.wisc.edu/real-time/westpac/winds/wgmsdlm1.GIF",
      "Satellate" :"https://www.cwa.gov.tw/V8/C/W/OBS_Sat.html?Area=2",
      "QPF" : "https://www.cwa.gov.tw/V8/C/P/QPF.html",
      "RainMap_accumulate" : f"http://tidpweather.vpnplus.to:50020/RainMap/Taiwan/TW_{now.year}{now.month}{now.day}_{now.hour}00.png",
      "RainMap_06" : f"http://tidpweather.vpnplus.to:50020/RainMap/Taiwan/TW_{now.year}{now.month}{now.day}_0600.png",
      "RainMap_yday" : f"http://tidpweather.vpnplus.to:50020/RainMap/Taiwan/TW_{now.year}{now.month}{now.day}_0000.png",
      "RainMap_oneHourAgo" : f"http://tidpweather.vpnplus.to:50020/RainMap/Taiwan/TW_{now.year}{now.month}{now.day}_{one_hours_ago.hour}00.png"
      }

# 設置webdriver頁面加載策略 (normal、eager、none)
options = Options()
options.page_load_strategy = "eager"

driver = webdriver.Chrome(options=options)

# 擷取氣象局圖資網址 (除了QPF)
def get_image_url(url, location, item):
    
    url_list = []
    
    try:
        driver.get(url)
        
        for i in location :
            element = driver.find_elements(By.CLASS_NAME, "img-responsive")[i]
            img_url = element.get_attribute("src")
            url_list.append(img_url)
            
    except Exception as e:
        print(f"{item} Exception:{e}")
        
    finally:
        return url_list


# 儲存成png檔案
def fetch_image(url, image_name):
    
    response = requests.get(url)
    
    if response.status_code == 200 :
        with Image.open(BytesIO(response.content)) as image:
            image.save(os.path.join(path, f"{image_name}.png"))
            
    elif url == URL["RainMap_accumulate"] and response.status_code != 200:
           print(f"RainMap has not been update cumulate rainfull for {now.hour}, using previous hours data")
           fetch_image(URL["RainMap_oneHourAgo"], "E_image")
           
    else :
         print(f"{image_name} Exception:{response.status_code}")


# 擷取QPF圖資網址
def QPF_fetch_image(QPF_url):
    
    QPF_list = []
    
    for url in QPF_url:
        pattern = r"_ChFcstPrecip_([\w_]+)\.png"
        match = re.search(pattern, url)
        img_name = match[1]
        
        try:
            response = requests.get(url)
            QPF_image = Image.open(BytesIO(response.content))
            QPF_image.save(os.path.join(path, f"{img_name}QPF.png"))
            QPF_list.append(QPF_image)
            
        except Exception as e:
            print(f"QPF Exception：{e}")
            
    return QPF_list


# 結合QPF圖資
def combine_image(QPF_list, forecast_range):
    
    bg = Image.new("RGBA", (2490, 3000), "#FFFFFF")
    
    for i in range(1,5):
        image = QPF_list[i-1]
        x = (i-1) % 2
        y = (i-1) // 2
        bg.paste(image, (x*1245, y*1500))
        
    bg.save(f"{forecast_range}hr_total.png")


# 地面天氣圖
SWM_url = get_image_url(URL["SWM"], [2], "地面天氣圖")
fetch_image(SWM_url[0], "SWM")

# 雷達迴波圖
Radar_url = get_image_url(URL["Radar"], [1], "雷達迴波圖")
fetch_image(Radar_url[0], "Radar")

# 流場圖
fetch_image(URL["StreamLine"], "StreamLine")

# 衛星雲圖
driver.get(URL["Satellate"])
button = driver.find_element(By.LINK_TEXT, "真實色")
button.click()
Satellite_Images_URL = driver.find_elements(By.CLASS_NAME, "img-responsive")[1].get_attribute("src")
fetch_image(Satellite_Images_URL, "Satellite")

# CWA定量降水
location_range = list(range(2,10)) # list中第2~10是QPF網址
QPF_url = get_image_url(URL["QPF"], location_range, "QPF")
QPF_list = QPF_fetch_image(QPF_url)
QPF_6hr = QPF_list[4:]
QPF_12hr = QPF_list[:4]
combine_image(QPF_6hr, 6)
combine_image(QPF_12hr, 12)

driver.quit()

# RainMap今日累積雨量圖資
fetch_image(URL["RainMap_accumulate"], "E_image")

# RainMap每日06時累積雨量
fetch_image(URL["RainMap_06"], "E_06_image")

# RainMap昨日累積雨量
fetch_image(URL["RainMap_yday"], "E_yday_image")