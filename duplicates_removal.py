# -*- coding: utf-8 -*-
"""
Created on Sat Sep  5 20:34:15 2020

@author: Ninja Xpress
"""

import pandas as pd

file = input("Copy and Paste the file name: ")

filename = (file+".xlsx")

print("Still loading, please wait...")

data = pd.read_excel(filename)

#To show the number of columns
n_column = data.columns.shape[0]

i = 0
k = 1

while i <= n_column and k <= n_column:
    
    #To display the number of unique and duplicated values
    x = data.iloc[:, i:k].duplicated().value_counts()
    x = x.rename(index = {False : "unique values", True : "duplicated"})
    
    #To show the name of each column
    print("Column name : ", data.iloc[:, i:k].columns[0], "\n", x, "\n")
    x = data.rename(index = {False : "unique values", True : "duplicated"})
    i += 1
    k += 1

nCol = data.columns.value_counts().sum() 
print(nCol," column(s) have been found!")

#Choose the column, then remove the duplicates    
select = str(input("Select the column: "))
data = data.drop_duplicates(subset = [select])

#Export to the new workbook    
data.to_excel(file+" (duplicates removed).xlsx", index = False)