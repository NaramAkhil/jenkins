'''
Created on Aug 13, 2020

@author: naram
'''


import xlrd
import xlwt
import string
varLocOfFile = "C:\\Users\\naram\\Downloads\\CheckList.xls"
rdwb = xlrd.open_workbook(varLocOfFile, formatting_info = True)
sheet = rdwb.sheet_by_index(0)
varNumOfRows = sheet.nrows

print("number of rows are",varNumOfRows)

for row in range(1, varNumOfRows):
    varSerialNum = sheet.cell_value(row,0)
    varMake = sheet.cell_value(row,1)
    varModel = sheet.cell_value(row,2)
    varYear = sheet.cell_value(row, 3)
    varProvience = sheet.cell_value(row,4)
    varPrice = sheet.cell_value(row,5)
    
    row = int(input("Enter a number: "))
    if (row % 2) :
        print("num is even")
    else:
        print("num is Odd")
       