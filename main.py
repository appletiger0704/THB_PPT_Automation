# -*- coding: utf-8 -*-
"""
Created on Mon Nov 20 14:06:04 2023

@author: User
"""

import subprocess
import os

os.chdir(r"C:\Users\user\Desktop\自動化separate\THB_PPT_Automation\program")

first_process = subprocess.Popen(["python", "Crawler.py"])
first_process.wait()



second_process = subprocess.Popen(["python", "circle_corner.py"])
second_process.wait()



thrid_process = subprocess.Popen(["python", "CWA_API.py"])
thrid_process.wait()



last_process = subprocess.Popen(["python", "PPTAutomation.py"])
last_process.wait()

