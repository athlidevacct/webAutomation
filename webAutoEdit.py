# -*- coding: utf-8 -*-
"""
Created on Sat Jun 27 10:07:25 2020

"""

from autoBooking import *
import time
import pandas as pd

webURL = ""

loginID = ""
loginPwd = ""
bookingCsv = "./editbooklist.csv"

bookdf = pd.read_csv(bookingCsv, delimiter=',')

autoBookObj = autoBooking(webURL, loginID, loginPwd)

autoBookObj.openBrowser()
autoBookObj.loginSystem()
time.sleep(2)
autoBookObj.closeToday()

currentWindowPage = autoBookObj.getCurrentWindow()

for row in bookdf.values:
    bookingID = row[0]
    startDate = row[1]
    startTime = row[2]
    endDate = row[3]
    endTime = row[4]
    
    autoBookObj.editBooking(bookingID,startDate,startTime,endDate,endTime)        
    autoBookObj.goWindows(currentWindowPage)
  
time.sleep(2)

autoBookObj.closeBrowser()
