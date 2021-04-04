# -*- coding: utf-8 -*-
"""
Created on Sun Apr  4 22:36:38 2021

@author: Admin
"""

import xlrd
import openpyxl
import easygui
import os

print("Please choose the folder which contains the files:\n")
filepath = easygui.diropenbox()

print("Please choose the folder which the result will be saved:\n")
pathsave = easygui.diropenbox()

tl = input("Please enter the row number of the title:\n")
tl = int(tl) - 1

g = os.walk(filepath)
path_xls = []

for path, dir_list, file_list in g:
    for file_name in file_list:
        path_xls.append(os.path.join(path, file_name))
        
data = []
wb_title = xlrd.open_workbook(path_xls[0])
data.append(wb_title.sheets()[0].row_values(tl))

for i in path_xls:
    wb = xlrd.open_workbook(i)
    for sheet in wb.sheets():
        for rownum in range(1,sheet.nrows):
            data.append(sheet.row_values(rownum))
            
wk = openpyxl.Workbook()

wkts = wk.active
for i in range(len(data)):
    for j in range(len(data[i])):
        wkts.cell(i+1,j+1,data[i][j])
pathsave = pathsave + "\\汇总.xlsx"
wk.save(pathsave)

input("Merge finished, press any button to exit:")