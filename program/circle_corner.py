# -*- coding: utf-8 -*-
"""
Created on Wed Nov 22 12:39:04 2023

@author: User
"""

from PIL import Image, ImageDraw, ImageOps
from datetime import datetime
import os


now = datetime.now()
now_str = now.strftime("%Y%m%d")

img_path = fr"C:\Users\user\Desktop\自動化separate\THB_PPT_Automation\{now_str}"

os.chdir(img_path)



def circle_corner(img, rad = 20, two_corner = True):
    
    # Image.new(mode, size, color) 用以創建圖像
    mask = Image.new("L", img.size, 0)
    
    # 將上述圖像設定為"可繪製圖像"
    draw = ImageDraw.Draw(mask)
    
    # 利用draw創造出的可繪製圖像，建立一個圓角的矩形
    draw.rounded_rectangle((0,0) + img.size, rad, fill = 255)
    
    # ImageOps 是一個可以對圖像進行操作的函示庫，其中fit是讓img的圖形大小符合mask圖形大小。
    # center指圖片擺放位置(0.5, 0.5)代表置中。
    round_img = ImageOps.fit(img, mask.size, centering = (0.5, 0.5))
    
    # putalpha 是一種Image方法，設定圖像的透明度
    round_img.putalpha(mask)
    
    return round_img



def apply_to_image(img_path):
    
    image = Image.open(img_path + ".png")
    
    # GIF圖檔本身帶有"調色版(palette-based)"的alpha通道，故與png本身支持透明度的格式有所衝突
    image = image.convert("RGBA")
    
    round_image = circle_corner(image)
    
    round_image.save("round_" + img_path + ".png", )
                  

path = ["SWM", "StreamLine", "Satellite"]

for i in path:
    apply_to_image(i)
    

