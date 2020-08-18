'''
Created on Aug 11, 2020

@author: naram
'''
from selenium import webdriver
import time
import xlrd
import xlwt
from xlutils.copy import copy


varLocOfFile = "C:\\Users\\naram\\OneDrive\\Documents\\test.xls"
rdwb = xlrd.open_workbook(varLocOfFile, formatting_info = True)
sheet = rdwb.sheet_by_index(0)
varNumOfRows = sheet.nrows
print("number of rows in sheet index0",varNumOfRows)

row = 1
while row < varNumOfRows:
    varTestCaseName = sheet.cell_value(row,1)
    varUserName = sheet.cell_value(row,2)
    varPassword = sheet.cell_value(row,3)
    varCreatedBy = sheet.cell_value(row,4)
    varActualRole = sheet.cell_value(row,5)
    varExceptedRole = sheet.cell_value(row, 6)
    print("row  num for index0 is ", row)
    print(varTestCaseName, varUserName, varPassword, varCreatedBy, varActualRole, varExceptedRole)
    if varUserName == "occ-ic":
        print("found successfully occ-ic")
        print("navigating to sheet index1")
        
        sheetIndexOne = rdwb.sheet_by_index(1)
        varNumOfRowsIndexOne = sheetIndexOne.nrows
        print("number of rows in sheet index1",varNumOfRowsIndexOne)
        for rowIndex1 in range(2, varNumOfRowsIndexOne):
            varSurName = sheetIndexOne.cell_value(rowIndex1, 1)
            varFirstName = sheetIndexOne.cell_value(rowIndex1, 2)
            varMiddleName = sheetIndexOne.cell_value(rowIndex1, 3)
            varSex = sheetIndexOne.cell_value(rowIndex1, 4)
            VarDob = sheetIndexOne.cell_value(rowIndex1, 5)
            varDod = sheetIndexOne.cell_value(rowIndex1, 6)
            print("row num for index1", rowIndex1)
            print(varSurName, varFirstName, varMiddleName,varSex, VarDob,varDod)
    else:
        print("not able find occ-ic")
    
    
        
    row = row + 1