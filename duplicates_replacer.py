# -*- coding: utf-8 -*-
"""
Created on Sat Sep  5 23:59:20 2020

@author: Ninja Xpress
"""
import pandas as pd

filename = input("Please input the file name : ")
col_name = input("Please input the column name : ")

data = pd.read_excel(filename+".xlsx")

repo = data[col_name].duplicated().value_counts()

#To rename values "False" as "Unique values" and "True" as "Duplicates" 
repo = repo.rename(index = {False : "Unique values ",
                    True : "Duplicates"})

print(repo)


repl = input("Input the desired value for the duplicates replacement :")

#To filter and show only the duplicated values 
dupl = data[col_name][data.duplicated([col_name])]

#To replace the new values into duplicated one
dupl = dupl.replace(dupl.values, repl)
data.loc[dupl.index, [col_name]] = dupl

#export to the new excel
data.to_excel(filename+" (modified).xlsx", index = False)
print("Successfully exported!")