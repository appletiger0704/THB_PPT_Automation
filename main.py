# -*- coding: utf-8 -*-
"""
Created on Mon Nov 20 14:06:04 2023

@author: User
"""

import subprocess

first_process = subprocess.Popen(["python", "Crawler.py"])
first_process.wait()

second_process = subprocess.Popen(["python", "PPTAutomation.py"])