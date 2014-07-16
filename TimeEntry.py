import datetime
import requests
from requests_ntlm import HttpNtlmAuth
from sharepoint import SharePointSite, basic_auth_opener
import urllib2
from ntlm import HTTPNtlmAuthHandler
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC # available since 2.26.0
import time
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
import Tkinter, tkFileDialog
import xlrd
import datetime
from Tkinter import *
import timeMenu
import timeMenu_support

def dataLoad():
    root = Tkinter.Tk()
    root.withdraw()
    file_path = tkFileDialog.askopenfilename(title = "Open Excel data file")
    workbook = xlrd.open_workbook(file_path)
    global filePathToServer
    filePathToServer = tkFileDialog.askopenfilename(title = "Open IEServer file")
    global sharepointData
    sharepointData = workbook.sheet_by_name('Sharepoint')
    global eliteData
    eliteData = workbook.sheet_by_name('Elite')
    global sharepointHoursPerDate
    sharepointHoursPerDate = {}
    for rowNum in range(1, sharepointData.nrows):
        if sharepointData.cell_value(rowNum, 1) != '' and sharepointData.cell_value(rowNum, 8) != '0169943-0001-MC' and datetime.datetime.strptime(sharepointData.cell_value(rowNum, 1), '%m/%d/%y').strftime('%m/%d/%y') in sharepointHoursPerDate:
            sharepointHoursPerDate[datetime.datetime.strptime(sharepointData.cell_value(rowNum, 1), '%m/%d/%y').strftime('%m/%d/%y')] += sharepointData.cell_value(rowNum, 3)
        elif sharepointData.cell_value(rowNum, 1) != '' and sharepointData.cell_value(rowNum, 8) != '0169943-0001-MC':
            sharepointHoursPerDate[datetime.datetime.strptime(sharepointData.cell_value(rowNum, 1), '%m/%d/%y').strftime('%m/%d/%y')] = sharepointData.cell_value(rowNum, 3)
    global startDate
    startDate = min(sharepointHoursPerDate.keys())

def SharePoint():
    driver = webdriver.Ie(filePathToServer)
    driver.get("https://tdem.getadvantage.com/Lists/Time%20Entry%20%20User%20Lookup/NewForm.aspx?RootFolder=%22%29%3B")
    # driver.get("https://tdemsandbox.getadvantage.com/Lists/Time%20Entry%20%20User%20Lookup/NewForm.aspx?RootFolder=")
    alert = driver.switch_to.alert
    alert.send_keys(timeMenu_support.PASSWORD)
    WebDriverWait(driver,45).until(EC.alert_is_not_present())
    time.sleep(2)
    for rowNum in range(1, sharepointData.nrows):
        if sharepointData.cell_value(rowNum, 1) != '' and sharepointData.cell_value(rowNum, 8) != '0169943-0001-MC':
            name = timeMenu_support.FULLNAME
            date = sharepointData.cell_value(rowNum, 1)
            program = sharepointData.cell_value(rowNum, 2)
            hours = str(sharepointData.cell_value(rowNum, 3))
            applicant = sharepointData.cell_value(rowNum, 4)
            pwnum = str(sharepointData.cell_value(rowNum, 5))
            pwnum = pwnum[:-2]
            task = sharepointData.cell_value(rowNum, 6)
            taskdesc = sharepointData.cell_value(rowNum, 7)

            nameInput = driver.find_element_by_id("Employee_d76b99b0-d0fc-432c-bcd7-03a848017a4a_$ClientPeoplePicker_EditorInput")
            nameInput.send_keys(name)

            dateInput = driver.find_element_by_id("Date_af5c6f10-4b1d-41f4-95cf-1bd599a23dc1_$DateTimeFieldDate")
            dateInput.send_keys(date)

            pa1791 = driver.find_element_by_id("Program_b1b6e7ad-ba2e-4d9a-8804-8f3aa5da2ae1_$RadioButtonChoiceField0")
            pa1791.click()

            taskSelect = Select(driver.find_element_by_id("Task_409b040c-9d88-4727-8374-6d1db4949ac3_$DropDownChoice"))
            taskSelect.select_by_visible_text(task)

            hoursInput = driver.find_element_by_id("Hours_x0020_Worked_e69f3e55-21eb-41e6-8711-0efa787b6229_$NumberField")
            hoursInput.send_keys(hours)

            applicantSelect = Select(driver.find_element_by_id("Applicant_6f089c8a-8a19-429b-9ba0-d5e1508fbac8_$DropDownChoice"))
            applicantSelect.select_by_visible_text(applicant)

            projectWorksheetInput = driver.find_element_by_id("Project_x0020_Worksheet_cbcee3af-dc93-4e0a-af5a-be88553ad806_$TextField")
            projectWorksheetInput.send_keys(pwnum)

            taskDescInput = driver.find_element_by_id("Title_fa564e0f-0c70-4ab9-b863-0177e6ddd247_$TextField")
            taskDescInput.send_keys(taskdesc)

            submitButton = driver.find_element_by_id("ctl00_ctl42_g_1e271c60_6ec2_4edb_adcc_f41e0142a685_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem")
            submitButton.click()
            time.sleep(3)
            driver.get("https://tdem.getadvantage.com/Lists/Time%20Entry%20%20User%20Lookup/NewForm.aspx?RootFolder=%22%29%3B")
            # driver.get("https://tdemsandbox.getadvantage.com/Lists/Time%20Entry%20%20User%20Lookup/NewForm.aspx?RootFolder=")

    time.sleep(4)
    driver.quit()

def Elite():
    driver = webdriver.Ie(filePathToServer)
    driver.get("http://elite.cohnreznick.net/webview/100Desktop/runtime/pgDisplayPage.aspx?pageno=105118&XSLPath=105TimeEntry/RunTime/pgTimeEntryFrame.xsl")
    driver.switch_to.frame("UPD1")
    for rowNum in range(1,eliteData.nrows):
        if eliteData.cell_value(rowNum, 3) > 0.00:
            name = timeMenu_support.FULLNAME
            date = eliteData.cell_value(rowNum, 1)
            directMatter = eliteData.cell_value(rowNum, 2)
            hours = str(eliteData.cell_value(rowNum, 3))
            narrative = eliteData.cell_value(rowNum, 4)
            serviceCode = eliteData.cell_value(rowNum, 5)
            directMatterInput = driver.find_element_by_name("timecard_dtmatter")
            directMatterInput.clear()
            directMatterInput.send_keys(directMatter)
            eliteDateInput = driver.find_element_by_name("dt_timecard_tworkdt1")
            eliteDateInput.clear()
            eliteDateInput.send_keys(date)
            eliteHoursInput = driver.find_element_by_name("nm_timecard_tworkhrs")
            eliteHoursInput.send_keys(hours)
            narrativeInput = driver.find_element_by_id("timedesc_tddesc")
            narrativeInput.send_keys(narrative)
            serviceCodeInput = driver.find_element_by_id("UserDef4")
            serviceCodeInput.send_keys(serviceCode)
            eliteSubmitButton = driver.find_element_by_id("btnAddUpdateButton")
            eliteSubmitButton.click()
            try:
                alert = browser.switch_to_alert()
                alert.dismiss
            except:
                pass
            time.sleep(1)
    time.sleep(2)
    driver.quit()

class TimeRecord(object):
    def __init__(self,date,hours,applicant,pwnum,task,taskdesc):
        """

        :rtype : object
        """
        self.date = date
        self.hours = hours
        self.applicant = applicant
        self.pwnum = pwnum
        self.task = task
        self.taskdesc = taskdesc
        if self.applicant == 'None':
            self.applicant = ''
    def __eq__(self, other):
        if self.date != other.date or self.hours != other.hours or self.applicant != other.applicant or self.pwnum != other.pwnum or self.task != other.task or self.taskdesc != other.taskdesc:
            return False
        else:
            return True

    def __ne__(self, other):
        return not self.__eq__(other)

    def __str__(self):
        return "Date: " + self.date + " " + "Hours: " + self.hours + " " + "Applicant: " + self.applicant + " " + "PW: " + self.pwnum + " " + "Task: " + self.task + " " + "Description: " + self.taskdesc

def sharepointCheck():
    print "Authenticating login"
    user = 'cr\\' + timeMenu_support.USER
    password = timeMenu_support.PASSWORD
    url = "https://tdem.getadvantage.com/"

    passman = urllib2.HTTPPasswordMgrWithDefaultRealm()
    passman.add_password(None, url, user, password)
    # create the NTLM authentication handler
    auth_NTLM = HTTPNtlmAuthHandler.HTTPNtlmAuthHandler(passman)
    # create and install the opener
    opener = urllib2.build_opener(auth_NTLM)
    urllib2.install_opener(opener)
    site = SharePointSite(url, opener)
    sp_list = site.lists['Time Entry - User Lookup']

    print "Retrieving data..."
    hoursPerDateEntered = {}
    for row in sp_list.rows:
        if row.Author['name'] == timeMenu_support.FULLNAME:
            # print(row.Author['name'], row.Hours_x0020_Worked, row.Applicant, row.Date.strftime('%m/%d/%Y'), row.Program)
            if row.Date.strftime('%m/%d/%y') in hoursPerDateEntered and row.Date.strftime('%m/%d/%y') >= startDate:
                hoursPerDateEntered[row.Date.strftime('%m/%d/%y')] += row.Hours_x0020_Worked
            elif row.Date.strftime('%m/%d/%y') >= startDate:
                hoursPerDateEntered[row.Date.strftime('%m/%d/%y')] = row.Hours_x0020_Worked
    print "Verifying Data..."
    for day in hoursPerDateEntered:
        hoursPerDateEntered[day] = round(hoursPerDateEntered[day],1)
    for day in sharepointHoursPerDate:
        sharepointHoursPerDate[day] = round(sharepointHoursPerDate[day],1)
    if hoursPerDateEntered == sharepointHoursPerDate:
        print "Everything matches up! You're done!"
    else:
        "Adding in missing data"
        listToAdd = []
        elementsOfDayFile = []
        elementsOfDaySharepoint = []
        daysMissed = []
        for day in sharepointHoursPerDate:
            if day not in hoursPerDateEntered or sharepointHoursPerDate[day] != hoursPerDateEntered[day]:
                daysMissed.append(day)
        for day in daysMissed:
            for rowNum in range(1, sharepointData.nrows):
                if sharepointData.cell_value(rowNum, 1) != '' and datetime.datetime.strptime(sharepointData.cell_value(rowNum, 1), '%m/%d/%y').strftime('%m/%d/%y') == day:
                    temp = TimeRecord(datetime.datetime.strptime(sharepointData.cell_value(rowNum, 1), '%m/%d/%y').strftime('%m/%d/%y'), str(sharepointData.cell_value(rowNum, 3)), sharepointData.cell_value(rowNum, 4), str(sharepointData.cell_value(rowNum, 5)), sharepointData.cell_value(rowNum, 6), sharepointData.cell_value(rowNum, 7))
                    elementsOfDayFile.append(temp)
        for row in sp_list.rows:
            if row.Author['name'] == timeMenu_support.FULLNAME and row.Date.strftime('%m/%d/%y') in daysMissed:
                temp = TimeRecord(row.Date.strftime('%m/%d/%y'),str(row.Hours_x0020_Worked), str(row.Applicant), str(row.Project_x0020_Worksheet), str(row.Task), str(row.Title))
                elementsOfDaySharepoint.append(temp)
        for item in elementsOfDayFile:
            if item not in elementsOfDaySharepoint:
                listToAdd.append(item)
        if len(listToAdd) > 0:
            addToSharepoint(listToAdd)

def addToSharepoint(recordList):
    driver = webdriver.Ie(filePathToServer)
    driver.get("https://tdem.getadvantage.com/Lists/Time%20Entry%20%20User%20Lookup/NewForm.aspx?RootFolder=%22%29%3B")
    # driver.get("https://tdemsandbox.getadvantage.com/Lists/Time%20Entry%20%20User%20Lookup/NewForm.aspx?RootFolder=")
    alert = driver.switch_to.alert
    alert.send_keys(timeMenu_support.PASSWORD)
    WebDriverWait(driver,45).until(EC.alert_is_not_present())
    time.sleep(2)
    for record in recordList:
        nameInput = driver.find_element_by_id("Employee_d76b99b0-d0fc-432c-bcd7-03a848017a4a_$ClientPeoplePicker_EditorInput")
        nameInput.send_keys(timeMenu_support.FULLNAME)

        dateInput = driver.find_element_by_id("Date_af5c6f10-4b1d-41f4-95cf-1bd599a23dc1_$DateTimeFieldDate")
        dateInput.send_keys(record.date)

        pa1791 = driver.find_element_by_id("Program_b1b6e7ad-ba2e-4d9a-8804-8f3aa5da2ae1_$RadioButtonChoiceField0")
        pa1791.click()

        taskSelect = Select(driver.find_element_by_id("Task_409b040c-9d88-4727-8374-6d1db4949ac3_$DropDownChoice"))
        taskSelect.select_by_visible_text(record.task)

        hoursInput = driver.find_element_by_id("Hours_x0020_Worked_e69f3e55-21eb-41e6-8711-0efa787b6229_$NumberField")
        hoursInput.send_keys(record.hours)

        applicantSelect = Select(driver.find_element_by_id("Applicant_6f089c8a-8a19-429b-9ba0-d5e1508fbac8_$DropDownChoice"))
        applicantSelect.select_by_visible_text(record.applicant)

        projectWorksheetInput = driver.find_element_by_id("Project_x0020_Worksheet_cbcee3af-dc93-4e0a-af5a-be88553ad806_$TextField")
        projectWorksheetInput.send_keys(record.pwnum)

        taskDescInput = driver.find_element_by_id("Title_fa564e0f-0c70-4ab9-b863-0177e6ddd247_$TextField")
        taskDescInput.send_keys(record.taskdesc)

        submitButton = driver.find_element_by_id("ctl00_ctl42_g_1e271c60_6ec2_4edb_adcc_f41e0142a685_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem")
        submitButton.click()
        time.sleep(2)
        driver.get("https://tdem.getadvantage.com/Lists/Time%20Entry%20%20User%20Lookup/NewForm.aspx?RootFolder=%22%29%3B")
        # driver.get("https://tdemsandbox.getadvantage.com/Lists/Time%20Entry%20%20User%20Lookup/NewForm.aspx?RootFolder=")

def main():
    timeMenu.vp_start_gui()
    try:
        if timeMenu_support.doCheck == True or timeMenu_support.doSharepoint == True or timeMenu_support.doElite == True:
            print "Loading data from file..."
            dataLoad()
            t0 = time.clock()
        if timeMenu_support.doElite == True:
            print "Starting Elite data upload...."
            Elite()
        if timeMenu_support.doSharepoint == True:
            print "Starting SharePoint data upload..."
            SharePoint()
        if timeMenu_support.doCheck == True:
            print "Starting SharePoint data verification..."
            sharepointCheck()
        print "Time elapsed: ", round(time.clock()-t0,1), "seconds"
        print "Exiting in 3 seconds..."
        time.sleep(5)
    except (AttributeError, UnboundLocalError):
        pass
    
main()