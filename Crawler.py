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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import re


now = datetime.now()
one_hours_ago = now - timedelta(hours = 1)
# yesterday = now - timedelta(days = 1)
# next_day = now + timedelta(days = 1)
# day_after_tomorrow = now + timedelta(days = 2)

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


options = Options()

# selenium headless模式，不顯示視窗
options.add_argument("--headless")

# 設置webdriver頁面加載策略 (normal、eager、none)
options.page_load_strategy = "normal"

driver = webdriver.Chrome(options=options)


def write_txt(text):
    
    date = now.strftime("%y%m%d")
    
    with open(os.path.join(path, f"{date}_img.txt"), 'a') as file:
        
        file.write(text + "\n")
        


# 擷取氣象局圖資網址 (除了QPF)
def get_image_url(url, location, item):
    
    url_list = []
    
    try:
        
        driver = webdriver.Chrome(options=options)
        driver.get(url)
        
        for i in location :
            
            # 先等瀏覽器載入所有元素，最多等10秒鐘
            elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "img-responsive")))
            
            img_url = elements[i].get_attribute("src")
            url_list.append(img_url)
            
    except Exception as e:
        
        write_txt(f"{item} Exception:{e}")
        
    finally:
        
        driver.quit()
        
        return url_list


# 儲存成png檔案ng檔案
def fetch_image(url, image_name):
    
    response = requests.get(url)
    
    if response.status_code == 200 :
        
        with Image.open(BytesIO(response.content)) as image:
            
            image.save(os.path.join(path, f"{image_name}.png"))
            
    elif url == URL["RainMap_accumulate"] and response.status_code != 200:
        
           write_txt(f"RainMap has not been update cumulate rainfull for {now.hour}, using previous hours data")
           print(f"RainMap has not been update cumulate rainfull for {now.hour}, using previous hours data")
           fetch_image(URL["RainMap_oneHourAgo"], "E_image")
           
    else :
        
         write_txt(f"{image_name} Exception:{response.status_code}" + "\n")
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
            
            write_txt(f"QPF Exception：{e}")
            
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
def SWM():
    
    SWM_url = get_image_url(URL["SWM"], [2], "地面天氣圖")
    fetch_image(SWM_url[0], "SWM")


# 雷達迴波圖
def Radar():
    
    Radar_url = get_image_url(URL["Radar"], [1], "雷達迴波圖")
    fetch_image(Radar_url[0], "Radar")


# 流場圖
def StreamLine():
    
    fetch_image(URL["StreamLine"], "StreamLine")


# 衛星雲圖
def Satellate():
    
    img_0700 = f"https://www.cwa.gov.tw/Data/satellite/TWI_VIS_TRGB_1375/TWI_VIS_TRGB_1375-{now.year}-{now.month}-{now.day}-07-00.jpg"

    if requests.get(img_0700).status_code == 200:
        
        fetch_image(img_0700, "Satellite")
        
    else:
        
        driver.get(URL["Satellate"])
        button = driver.find_element(By.LINK_TEXT, "真實色")
        button.click()
        Satellite_Images_URL = driver.find_elements(By.CLASS_NAME, "img-responsive")[1].get_attribute("src")
        fetch_image(Satellite_Images_URL, "Satellite")


# CWA定量降水
def QPF():
    
    location_range = list(range(2,10)) # list中第2~10是QPF網址
    QPF_url = get_image_url(URL["QPF"], location_range, "QPF")
    QPF_list = QPF_fetch_image(QPF_url)
    QPF_6hr = QPF_list[4:]
    QPF_12hr = QPF_list[:4]
    combine_image(QPF_6hr, 6)
    combine_image(QPF_12hr, 12)


def RainMap():
    
    # RainMap今日累積雨量圖資
    fetch_image(URL["RainMap_accumulate"], "E_image")
    
    # RainMap每日06時累積雨量
    fetch_image(URL["RainMap_06"], "E_06_image")
    
    # RainMap昨日累積雨量
    fetch_image(URL["RainMap_yday"], "E_yday_image")



# if __name__ == "__main__":
#     p1 = multiprocessing.Process(target=SWM)
#     p2 = multiprocessing.Process(target=Radar)
#     p3 = multiprocessing.Process(target=StreamLine)
#     p4 = multiprocessing.Process(target=Satellate)
#     p5 = multiprocessing.Process(target=QPF)
#     # p6 = multiprocessing.Process(target=RainMap)
    
#     p1.start()
#     p2.start()
#     p3.start()
#     p4.start()
#     p5.start()
#     # p6.start()
    
#     p1.join()
#     p2.join()
#     p3.join()
#     p4.join()
#     p5.join()
#     # p6.join()


thread1 = threading.Thread(target = SWM)
thread2 = threading.Thread(target = Radar)
thread3 = threading.Thread(target = StreamLine)
thread4 = threading.Thread(target = Satellate)
thread5 = threading.Thread(target = QPF)
thread6 = threading.Thread(target = RainMap)
        
thread1.start()
thread2.start()
thread3.start()
thread4.start()
thread5.start()
thread6.start()

thread1.join()
thread2.join()
thread3.join()
thread4.join()
thread5.join()
thread6.join()

driver.quit()