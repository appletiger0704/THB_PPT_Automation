# -*- coding: utf-8 -*-
"""
Created on Mon Nov 20 14:06:04 2023

@author: User
"""

import subprocess
import time


start = time.time()
first_process = subprocess.Popen(["python", "Crawler.py"])
first_process.wait()

second_process = subprocess.Popen(["python", "circle_corner.py"])
second_process.wait()

last_process = subprocess.Popen(["python", "PPTAutomation.py"])

end = time.time()
cost_time = end - start
print(f"程式花費時間：{cost_time}")

