# -*- coding:utf-8 -*-
#History:
#2015/10/14 hqy version 3.0
#Any question contact : 302988766@qq.com
#

import xlwings
from os import listdir, getcwd, mkdir, path, chdir
import re, shutil



#set default variables
currentpath = getcwd()
resultpath = "%s\\result" % getcwd()

#make result dictionary
if not path.isdir(resultpath):
    mkdir(resultpath)

#get files list and copy source data to Template
files = listdir(getcwd())
wbT = xlwings.Workbook(r"%s\Template.xls" % currentpath)

for file in files:
    if file.endswith(".xls") and file != "Template.xls":
        wbA = xlwings.Workbook(r"%s\%s" % (currentpath, file))
        xlwings.Range("A26:M34", wkb = wbT).value = xlwings.Range("A12:M20",wkb = wbA).value
        wbA.close()      
                
        wbT.save(r"%s\%s" % (currentpath,re.sub("(.+?).xls", r"\1_result.xls" , file)))
wbT.close()

#move results to dictionary: result.
files = listdir(getcwd())
for file in files:
    if file.find("_result") != -1 :
        print(file.find("_result"))
        shutil.move(r"%s\%s" % (currentpath,file) , r"%s\%s" % (resultpath,file))
        
#create a Result.xlsx and write summary to Result.xlsx file 
chdir(resultpath)
files = listdir(getcwd())
wbR = xlwings.Workbook()
count = 1
for file in files:
    
    if file.endswith(".xls") and file != "Result.xls":
        print(file)
        wbA = xlwings.Workbook(r"%s\%s" % (resultpath,file))
        xlwings.Range("A%s" %  count , wkb = wbR).value = file
    
        xlwings.Range("sheet1","B%s:L%s" % (count, count+5), wkb = wbR).value = xlwings.Range("sheet1","N16:X21", wbk = wbA).value
        
        print(xlwings.Range("B%s:L%s" % (count, count+5), wkb = wbR).value)
        wbA.close()
        count += 6
        
#format output
xlwings.Range("A1:L%s" % str(count-1), wkb = wbR).number_format = "0.0" #keep only one decimal digit
xlwings.Range("A1:L%s" % str(count-1), wkb = wbR).autofit("c")    #autofit columns

wbR.save(r"%s\Result.xlsx" % resultpath)



