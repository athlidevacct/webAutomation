# -*- coding: utf-8 -*-
"""
Created on Sat Jun 27 09:34:15 2020

"""
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
from datetime import datetime

import time

class autoBooking:
    
    def __init__(self, loginUrl, userID, userPwd):
        self.driverPath = ".\chromedriver.exe"
        self.webURL = loginUrl
        self.loginID = userID
        self.loginPwd = userPwd
        self.bookingList = None
        self.driver = webdriver.Chrome(self.driverPath)
        
    def diffMonth(month):
        today = datetime.today()
        return int(month) - int(today.month)

    def splittimestampe(dateString):
        datePart = dateString.split(" ")
        
        month = datePart[1].split("/")[1]
        date = datePart[1].split("/")[0]
        hour = datePart[2].split(":")[0]
        mins = datePart[2].split(":")[1]
        if mins == "00":
            mins = "0"
        time = str(int(hour)) + ":" + mins
        return month,date,time
    
    def splittimestampe2(timeString):
       
        hour = timeString.split(":")[0]
        mins = timeString.split(":")[1]
        if mins == "00":
            mins = "0"
        time = str(int(hour)) + ":" + mins
        return time
    
    def splitDate(dateString):
       
        month = dateString.split("/")[1]
        date = dateString.split("/")[0]

        return month, date
    
    def openBrowser(self):
        return self.driver.get(self.webURL)
    
    def closeBrowser(self):
        self.driver.close()
        
    def loginSystem(self):
        input_id = self.driver.find_element_by_id("txtUserName")
        input_id.send_keys(self.loginID)
        
        input_pwd = self.driver.find_element_by_id("txtPassword")
        input_pwd.send_keys(self.loginPwd)
        
        self.driver.find_element_by_id("btnLogin").click()
    
    def closeToday(self):
        closeToday = self.driver.find_elements_by_link_text("Close Today Page")
        for i in closeToday:
            i.click()
            
    def goWindows(self, windowName):
        self.driver.switch_to.window(windowName)
        
    def getCurrentWindow(self):
        return self.driver.current_window_handle

    def switchFrame(self, frameName):
        self.driver.switch_to.frame(frameName)
        
    def getElement(self, elementName):
        return self.driver.find_element_by_name(elementName)
    
    def getLink(self, elementName):
        return self.driver.find_element_by_link_text(elementName)
    
    def editBooking(self, bookingID, startDate, startTime, endDate, endTime):
        wait = WebDriverWait(self.driver, 10000)
        isFound = autoBooking.yourBooking(self, bookingID)
        if isFound:   
            self.driver.find_element_by_xpath('//*[@id="tblRoomBookingGrid"]/tbody/tr[2]/td[12]/input[1]').click()
            bookingFormMain = wait.until(EC.presence_of_element_located((By.ID, 'BookingFormMainTable')))
            bookingFormMain.find_element_by_id('edittimeFrmTxtLink').click()
            timeFrmEditShow = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'timeFrmEditShow')))
            
            #Select Date
            timeFrmEditShow.find_element_by_id("StartDate").click()           
            wait.until(EC.visibility_of_element_located((By.ID, 'calendarFrame')))
            self.driver.switch_to.frame("calendarFrame")
            #shift the month
            smonth, sdate = autoBooking.splitDate(startDate)
            emonth, edate = autoBooking.splitDate(endDate)
            
            for d in range(autoBooking.diffMonth(smonth)):
                nextMonthPath =  self.driver.find_element_by_xpath('//*[@id="Calendar1"]/tbody/tr[1]/td/table/tbody/tr/td[3]/a')
                # Make click in that button
                ActionChains(self.driver).click(nextMonthPath).perform()
                time.sleep(1)
            #click the date
            datePicker_s = self.driver.find_element_by_link_text(str(sdate))
            datePicker_s.click()
     
            self.driver.switch_to.parent_frame()
      
            #select time
            Select(timeFrmEditShow.find_element_by_name("startTime")).select_by_value(startTime)
            Select(timeFrmEditShow.find_element_by_name("endTime")).select_by_value(endTime)
                      
            timeFrmEditShow.find_element_by_id('btn_EditTime').click()
            self.driver.find_element_by_id('saveAndClose').click()
            time.sleep(1)
            self.driver.find_element_by_id('btnSendEmail').click()
            time.sleep(1)


   
        
    def yourBooking(self, bookingID):
        wait = WebDriverWait(self.driver, 10)
        wait.until(EC.presence_of_element_located((By.ID, 'leftNavigation')))
        self.driver.switch_to.frame("leftNavigation")
        #Find a room
        findroom = self.driver.find_elements_by_link_text("Your Bookings")
        for f in findroom:
            f.click()
        
        self.driver.switch_to.parent_frame()#.frame("mainDisplayFrame")#.parent_frame()
        time.sleep(2)
        self.driver.switch_to.frame("mainDisplayFrame")
        
        time.sleep(1)
       
        referenceID = self.driver.find_element_by_id('referenceID')
        referenceID.send_keys(bookingID)
        
        self.driver.find_element_by_name('btnFilter').click()
        
        table = wait.until(EC.presence_of_element_located((By.ID, 'tblRoomBookingGrid')))
        rows = table.find_elements_by_xpath('//*[@id="tblRoomBookingGrid"]/tbody/tr')
        num_columns = len(self.driver.find_elements_by_xpath('//*[@id="tblRoomBookingGrid"]/tbody/tr/td'))
        
        if num_columns <= 1:
            notFound = table.find_element_by_xpath('//*[@id="tblRoomBookingGrid"]/tbody/tr[2]/td').text
            print(notFound)
            return False
        else:
            return True
            
        
            
    def deleteBooking(self, bookingID, currentWindowPage):
        wait = WebDriverWait(self.driver, 10000)
        isFound = autoBooking.yourBooking(self, bookingID)
        if isFound:   
            self.driver.find_element_by_xpath('//*[@id="tblRoomBookingGrid"]/tbody/tr[2]/td[12]/input[2]').click()
            time.sleep(5)
            for handle in self.driver.window_handles:
                if handle != currentWindowPage: 
                    delete_page = handle
                    self.driver.switch_to.window(delete_page)
                    #Click Save button
                    self.driver.find_element_by_xpath('//*[@id="tblMainTable"]/tbody/tr[2]/td/input[1]').click()
                    time.sleep(1)
            
        
        
    def searchRoom(self, location, startDate, endDate):
        
        wait = WebDriverWait(self.driver, 10)
        
        smonth,sdate,starttime = autoBooking.splittimestampe(startDate)
        emonth,edate,endtime = autoBooking.splittimestampe(endDate)
    
        wait.until(EC.presence_of_element_located((By.ID, 'leftNavigation')))
        self.driver.switch_to.frame("leftNavigation")
        #Find a room
        findroom = self.driver.find_elements_by_link_text("Find a Room")
        for f in findroom:
            f.click()
         
        self.driver.switch_to.parent_frame()#.frame("mainDisplayFrame")#.parent_frame()
        time.sleep(2)
        self.driver.switch_to.frame("mainDisplayFrame")

        searchOption = wait.until(EC.visibility_of_element_located((By.ID, 'searchOptions')))
        
        groupID = Select(searchOption.find_element_by_name("groupID"))
        #groupID = Select(self.driver.find_element_by_name("groupID"))
        groupID.deselect_all()
        groupID.select_by_value('43')
        
        roomName = self.driver.find_element_by_id("resourceName")
        roomName.send_keys(location)
        
        #select time
        starTime = Select(self.driver.find_element_by_name("startTime"))
        starTime.select_by_value(starttime)
        
        endTime = Select(self.driver.find_element_by_name("endTime"))
        endTime.select_by_value(endtime)
    
        #Select Date
        startDate = self.driver.find_element_by_id("StartDate")
        startDate.click()
        
        wait.until(EC.visibility_of_element_located((By.ID, 'calendarFrame')))
        self.driver.switch_to.frame("calendarFrame")
        #shift the month
        for d in range(autoBooking.diffMonth(emonth)):
            nextMonthPath =  self.driver.find_element_by_xpath('//*[@id="Calendar1"]/tbody/tr[1]/td/table/tbody/tr/td[3]/a')
            # Make click in that button
            ActionChains(self.driver).click(nextMonthPath).perform()
            time.sleep(1)
        
        #click the date
        datePicker_s = self.driver.find_element_by_link_text(str(edate))
        #datePicker_s = driver.find_element_by_xpath('//*[@id="Calendar1"]/tbody/tr[7]/td[3]/a')
        datePicker_s.click()
 
        self.driver.switch_to.parent_frame()
        
        #driver.close()
        self.driver.find_element_by_id("Find").click()
        
        try:
            wait.until(EC.presence_of_element_located((By.ID, 'searchResults')))
        finally:
            return True
        
        

    def bookRoom(self, subject, hostedBy, currentWindowPage):
        wait = WebDriverWait(self.driver, 10)
        outerSearchResults = wait.until(EC.presence_of_element_located((By.ID, 'OuterSearchResults')))
        bookBtn = outerSearchResults.find_element_by_xpath('//*[@id="searchResultsRepeater_ctl01_findActionButton"]')
        bookBtn.click()
        
        time.sleep(5)
        for handle in self.driver.window_handles:
            if handle != currentWindowPage: 
                booking_page = handle
                self.driver.switch_to.window(booking_page)
                title = self.driver.find_element_by_id('gen_meetingTitle')
                title.send_keys(subject)
                hostname = self.driver.find_element_by_id('gen_Host')
                hostname.send_keys(hostedBy)
                time.sleep(1)
                hostname.send_keys(Keys.ARROW_DOWN)
                #wait for first dropdown option
                time.sleep(1)
                #first_option = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "ui-menu-item-wrapper")))
                #first_option = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.CLASS_NAME, "ui-menu-item-wrapper")))
                hostname.send_keys(Keys.RETURN)
                #Click Save button
                self.driver.find_element_by_id('saveAndClose').click()
                time.sleep(1)
                self.driver.find_element_by_id('btnSendEmail').click()
                time.sleep(1)
            
    

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
        
