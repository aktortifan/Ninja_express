#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  9 22:46:35 2020

@author: AktorEvan
"""

import pandas as pd

doc = input("Please input the file name: ")
colName = input ("Please input the column name: ")

print("Loading the file, please wait..", "\n")

file = pd.ExcelFile(doc+".xlsx")
sheet_list = file.sheet_names

print(len(file.sheet_names), "sheet(s) found in the file named", doc, "\n")
i = 0
k = 1

listSheet = []
listDupl = []

while i < len(sheet_list):
    
    #List for dataframe
    listSheet.append(i)
    listSheet[i] = pd.read_excel(doc+".xlsx", sheet_name=sheet_list[i])
    

    #To display the number of unique and duplicated values
    x = listSheet[i].duplicated(colName).value_counts()
    x = x.rename(index = {False : "Unique values", True : " Duplicates"})
     
    #To show the name of each sheet
    print("Sheet name : ", sheet_list[i], "(", colName, ")", "\n", x, "\n")
   
    
    #list for duplicates
    listDupl.append(i)
    listDupl[i] = listSheet[i][listSheet[i].duplicated(colName)]
    

    #Export duplicates to .xlsx
    listDupl[i].to_excel(sheet_list[i]+" (duplicate lists).xlsx", sheet_name = sheet_list[i], index = False)
    i += 1
    k += 1 
    
print("Succefully scanned!")