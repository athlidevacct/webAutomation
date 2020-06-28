# -*- coding: utf-8 -*-
"""
Created on Sun Jun 28 15:52:33 2020

@author: alfredli
"""
from autoBooking import *
import time
import pandas as pd

webURL = ""

loginID = ""
loginPwd = ""
bookingCsv = "./deletebooklist.csv"

bookdf = pd.read_csv(bookingCsv, delimiter=',')

autoBookObj = autoBooking(webURL, loginID, loginPwd)

autoBookObj.openBrowser()
autoBookObj.loginSystem()
time.sleep(2)
autoBookObj.closeToday()

currentWindowPage = autoBookObj.getCurrentWindow()
    
for row in bookdf.values:
    bookingID = row[0]
    print(bookingID)
    autoBookObj.deleteBooking(int(bookingID), currentWindowPage)        
    autoBookObj.goWindows(currentWindowPage)

#autoBookObj.goWindows(currentWindowPage) 
time.sleep(2)

autoBookObj.closeBrowser()