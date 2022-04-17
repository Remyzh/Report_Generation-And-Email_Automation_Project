import os
import time
import itertools
from win32com.client import Dispatch

listOfFileNames = []
listOfSheetNames = []
path_definitions = 'C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\Definitions.xlsx'    #VARIABLE: NEED TO CHANGE


def getSheetNames_function():
    os.chdir("C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test")       #VARIABLE: NEED TO CHANGE

    a = open("File_names.txt", "w")

    for path, subdirs, files in os.walk(r'C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\AccountManagers'):  #VARIABLE: NEED TO CHANGE
        for filename in files:
            a.write(str(filename))
            a.write("\n")

def getFileNames_function():
    os.chdir("C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test")       #VARIABLE: NEED TO CHANGE

    a = open("Sheet_names.txt", "w")

    for path, subdirs, files in os.walk(r'C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\AccountManagers'):     #VARIABLE: NEED TO CHANGE
        for filename in files:
            file_name, file_type = os.path.splitext(filename)
            
            ManagerName,file_info = file_name.split("'")
            a.write(str(ManagerName) + str(" Results"))
            a.write("\n")

        

getSheetNames_function()
getFileNames_function()

#Stores file names from txt file as a list data type in python
with open("File_names.txt", 'r') as f:                      
    listOfFileNames = [line.rstrip() for line in f] 
    print(listOfFileNames)

#Stores converted kathy workbook tab format from txt file as a list data type in python
with open("Sheet_names.txt", 'r') as f:                     
    listOfSheetNames = [line.rstrip() for line in f] 
    #print(listOfSheetNames)


#This is the main loop
for workbookName, worksheetName in zip(listOfFileNames, listOfSheetNames):
    path1 = 'C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\source.xlsx'            #VARIABLE: NEED TO CHANGE 
    # path of the source file VARIABLE ^^
    link2 = 'C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\AccountManagers\\' + str(workbookName)     #VARIABLE: NEED TO CHANGE   
    # path of the destination file VARIABLE ^^
    path2 = link2.rstrip()
    print(path2)
    
    xl = Dispatch("Excel.Application")
    xl.Visible = False     #You can remove this line if you don't want the Excel application to be visible
    xl.Interactive = False

    wb1 = xl.Workbooks.Open(Filename=path1)
    wb2 = xl.Workbooks.Open(Filename=path2)
    wb3 = xl.Workbooks.Open(Filename=path_definitions)

    ws1 = wb1.Worksheets(str(worksheetName))        #put sheetname that you want to copy (the persons name + format) VARIABLE
    ws1.Copy(Before=wb2.Worksheets(1))              #set position in destination file you want to paste in. VARIABLE
    ws2 = wb3.WorkSheets("Definitions")
    ws2.Copy(Before=wb2.Worksheets(2))


    wb2.Close(SaveChanges=True)
    xl.Quit()
    time.sleep(0.7)

