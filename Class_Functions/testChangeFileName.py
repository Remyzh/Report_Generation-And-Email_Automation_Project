#Please note this code is not complete yet!

#Purpose: When the code runs, it updates all the attachment file names in the first col of every CSV file with the proper reporting period file name.

import pandas as pd
import datetime
import os
import pathlib
import win32com.client
import csv

#Below is the code to get today's date, so we can tell the month.
dateTime = datetime.datetime.today()

def fileName_converter(fileName):

    strEmail = "".join(fileName)
    newFileName = strEmail.split("'s")[0] #Not sure if this works. I am trying to split the file name to before and after the "'s" which should in theory give you the first half, which is the name of the AM.
    return(newFileName) #Then after returning the name of the Am with "'s" you can add the proper reporting period + " results"

with open("email_list.txt", 'r') as f:                      
    listOfEmail = [line.rstrip() for line in f] 

#Below is the dictionary containing the response for each month. We need this because the datetime function outputs a number for the month instead of the actually name of each month.
month_Quarter = {1:'Q1 FY22', 2:'February FY22', 3:'March FY22', 4:'Q2 FY22', 5:'May FY22', 6:'June FY22', 7:'Q3 FY22', 8:'August FY22', 9: 'September FY22', 10:'Q4 FY22', 11:'November FY22', 12:'December FY22'}

#------------------------------------------------------------------------------------------------------------------------#


for csvFile in listOfEmail:
##Main Loop: One loop represents one email, the loop ends when it is done going through all csv/excel Files.
    # open the file in read mode
    filename = open('EmailLists\\' + str(csvFile), 'w')
    
    # creating dictreader object
    file = csv.writer(filename)
    for col in file:
        #This secondary loop is to loop through ONE csv file
        #This loop is incomplete,
        file.writerow["File Names"] = fileName_converter(col) + str(month_Quarter[dateTime.month-1]) + " results"
