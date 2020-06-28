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
bookingCsv = "./booklist.csv"

bookdf = pd.read_csv(bookingCsv, delimiter=',')

autoBookObj = autoBooking(webURL, loginID, loginPwd)

autoBookObj.openBrowser()
autoBookObj.loginSystem()
time.sleep(2)
autoBookObj.closeToday()

currentWindowPage = autoBookObj.getCurrentWindow()

for row in bookdf.values:
    subject = row[0]
    location = row[1]
    hostedBy = row[2]
    hostedEmail = row[3]
    startDate = row[4]
    endDate = row[5]
    
    isFind = autoBookObj.searchRoom(location, startDate, endDate)
    if isFind == True:
        autoBookObj.bookRoom(subject, hostedBy, currentWindowPage)
        
    autoBookObj.goWindows(currentWindowPage)
    
time.sleep(2)

autoBookObj.closeBrowser()
