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

class p4Record(object):
    def __init__(self, pwnum, date, amount, signdate):
        """

        :rtype : object
        """
        self.date = datetime.datetime.strptime(date, '%m-%d-%y')
        self.amount = str(amount)
        self.pwnum = str(pwnum)
        self.signdate = datetime.datetime.strptime(signdate, '%m-%d-%y')
        self.rownum = -1
    def __eq__(self, other):
        if self.date != other.date or self.amount != other.amount or self.pwnum != other.pwnum or self.signdate != other.signdate:
            return False
        else:
            return True

    def __ne__(self, other):
        return not self.__eq__(other)

    def __str__(self):
        return  "Amount: " + self.amount + " " + "PW: " + self.pwnum + " " + "RowNum: " + str(self.rownum)

file_path = "C:\Users\schitalia\Documents\P4Updates.xlsx"
workbook = xlrd.open_workbook(file_path)
p4Data = workbook.sheet_by_name('toUpdate')
p4ToAdd = []
for rowNum in range(1, p4Data.nrows):
    temp = p4Record(p4Data.cell_value(rowNum,0),p4Data.cell_value(rowNum,1),p4Data.cell_value(rowNum,2),p4Data.cell_value(rowNum,3))
    p4ToAdd.append(temp)

user = 'cr\schitalia'
password = 'sharepointpassword'
url = "https://tdem.getadvantage.com/"

passman = urllib2.HTTPPasswordMgrWithDefaultRealm()
passman.add_password(None, url, user, password)
auth_NTLM = HTTPNtlmAuthHandler.HTTPNtlmAuthHandler(passman)
opener = urllib2.build_opener(auth_NTLM)
urllib2.install_opener(opener)
site = SharePointSite(url, opener)
sp_list = site.lists['Project Status']
projectList = sp_list.rows
suffix = '-PA-1791'
for item in p4ToAdd:
    fullItem = item.pwnum + suffix
    for x in range(len(projectList)):
        if projectList[x].name == fullItem:
            item.rownum = x
            break

for item in p4ToAdd:
    print item
    sp_list.rows[item.rownum].P_x002e_4_x003a__x0020_Amount_x0 = item.amount
    sp_list.rows[item.rownum].P_x002e_4_x003a__x0020_Actual_x0 = item.date
    sp_list.rows[item.rownum].P_x002e_4_x003a__x0020_Date_x002 = item.signdate

sp_list.save()
print "done, uploaded", len(p4ToAdd), "items"