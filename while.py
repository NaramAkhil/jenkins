'''
Created on Aug 4, 2020

@author: naram
'''
from selenium import webdriver
import time
import xlrd
import xlwt
from xlutils.copy import copy


varLocOfFile = "C:\\Users\\naram\\OneDrive\\Documents\\TestScenariosV3.xls"
rdwb = xlrd.open_workbook(varLocOfFile, formatting_info = True)
sheet = rdwb.sheet_by_index(0)
varNumOfRows = sheet.nrows
print("number of rows are",varNumOfRows)


wtwb = copy(rdwb)
#write arg will take 3 arg
#wtwb.get_sheet(0).write(2,9,"Test")

row = 2
while row < varNumOfRows:
    if sheet.cell_value(row, 7).upper() == "Y":
        
        varTestCaseName = sheet.cell_value(row,1)
        varUserName = sheet.cell_value(row,3)
        varPassword = sheet.cell_value(row,4)
        varCustomerName = sheet.cell_value(row,5)
        varExceptedValue = sheet.cell_value(row,6)
        
        
        br = webdriver.Chrome()
        br.get("https://stage.quinc.mcscs.jus.gov.on.ca/Pages/C/QuinC/HomeQuinC.aspx")
        elmId = br.find_element_by_id("LoginUser_UserName").send_keys(varUserName)
        elmId = br.find_element_by_id("LoginUser_Password").send_keys(varPassword)
        elmBtn = br.find_element_by_xpath('//*[@id="LoginUser_btnLogin"]').click()
        time.sleep(5)
        
        elmRow = br.find_element_by_partial_link_text(varCustomerName).click()
        time.sleep(3)
        varActualCreatedBy = br.find_element_by_id('ctl00_MainContent_ApplicantDetailUserControl_lblCreateUser').text
        wtwb.get_sheet(0).write(row,8,varActualCreatedBy)
        if(varActualCreatedBy == varExceptedValue):
            varResult = "pass" 
            print("Test Case",varTestCaseName, "ended as ", varResult)
            wtwb.get_sheet(0).write(row,9,varResult)
        else:
            varResult = "fail"
            print("Test Case", varTestCaseName, "ended as", varResult)
            wtwb.get_sheet(0).write(row,9,varResult)
        elmLogout = br.find_element_by_partial_link_text('Log Out').click()
        time.sleep(3)
        br.quit()
    else:
        varResult = "Skipping test case as execute flag"
        print("Skipping test case as execute flag is set to N")
        wtwb.get_sheet(0).write(row,9,varResult)
    row = row + 1

wtwb.save ("C:\\Users\\naram\\Downloads\\output.xls")
