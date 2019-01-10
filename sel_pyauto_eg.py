#!/usr/bin/python

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import sys
import unittest
import os
import getpass
import time
import pyautogui
import shutil
import pandas as pd

def presskey(keyname,num,delay):

    for n in range(0,num):
        pyautogui.press(keyname)
        time.sleep(delay)
		
df = pd.read_excel("input.xlsx")
k = df['County'].values.T.tolist()
v = df['Sequence'].values.T.tolist()

driver = webdriver.Chrome(executable_path=os.path.join(os.getcwd(),'drivers\\chromedriver.exe'))   

#driver.get("https://cogcc.state.co.us/production/?&apiCounty=123&apiSequence=38905")

def get_url():

    for i in range(0,len(k)):
    
        api = str(k[i])
        seq = str(v[i])		
		
        url = "https://cogcc.state.co.us/production/?&apiCounty=" + api + "&apiSequence=" + seq

        driver.get(url)
        presskey('tab',4,2)
        presskey('enter',1,2)
        presskey('down',1,2)
        time.sleep(4)
        presskey('enter',1,2)

        shutil.move("C:\\Users\\Ritesh Sharma\\Downloads\\flatProduction.csv", "D:\\upwork_10\\" + api + "_" + seq + ".csv")

get_url()

driver.close()