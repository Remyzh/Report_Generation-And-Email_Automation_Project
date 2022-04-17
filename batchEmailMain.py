## Below are all the Python libraries that we need to import
import pandas as pd
import datetime
import os
import pathlib
import win32com.client
import csv
import re

outlook = win32com.client.Dispatch('outlook.application') #This is a a definition so we can save time
startBold = '\033[1m'
endBold = "\033[0m"
makeBlue = '\u001b[34m'
#------------------------------------------------------------------------------------------------------------------------#

## Below are all the individual functions, most of them for converting formats.

#The convert() function adds ";" to the string, this is used to seperate each of the email addresses inputed into outlook
def convert(s):
  
    # initialization of string to ""
    new = ""
  
    # traverse in the string 
    for x in s:
        new += x
        new += '; ' 
  
    # return string 
    return new

#The email_to_name_converter() function inputs an RBC email and outputs the first name capitlized of the email owner. 
def email_to_name_converter(email):
    strEmail = "".join(email)

    name, _, _ = strEmail.partition(".") 

    return(name.capitalize())

#The email_to_Fullname_converter() function inputs an RBC email and outputs the full name of the email owner.
def email_to_Fullname_converter(email):
    strEmail = "".join(email)
    Fullname = strEmail.split("@")[0]
    return(Fullname)

#the getSheetNames_function() function creates a txt file email_list.txt containing the names of all the emails
def getSheetNames_function():
    os.chdir("C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test")       #VARIABLE: NEED TO CHANGE, this is the location of the overall folder

    a = open("email_list.txt", "w")

    for path, subdirs, files in os.walk(r'C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\EmailLists'):  #VARIABLE: NEED TO CHANGE, change this to location of Email List
        for filename in files:
            a.write(str(filename))
            a.write("\n")

#Below is the dictionary containing the response for each month. We need this because the datetime function outputs a number for the month instead of the actually name of each month.
month_Quarter = {1:'Q1 FY22', 2:'February FY22', 3:'March FY22', 4:'Q2 FY22', 5:'May FY22', 6:'June FY22', 7:'Q3 FY22', 8:'August FY22', 9: 'September FY22', 10:'Q4 FY22', 11:'November FY22', 12:'December FY22'}

#Below is the code to get today's date, so we can tell the month.
dateTime = datetime.datetime.today()

getSheetNames_function()    # This is so we can run the getSheetNames_function()

with open("email_list.txt", 'r') as f:                      
    listOfEmail = [line.rstrip() for line in f] 

#------------------------------------------------------------------------------------------------------------------------#

##Main Loop: One loop represents one email, the loop ends when it is done going through all lines of a csvFile.

for csvFile in listOfEmail:
    # open the file in read mode
    filename = open('EmailLists\\' + str(csvFile), 'r')
    
    # creating dictreader object
    file = csv.DictReader(filename)
    
    # creating empty lists
    file_names= []
    address_to = []
    cc_to = []
    
    # iterating over each row and append
    # values to empty list
    for col in file:
        file_names.append(col["File Name"])          # + '.xlsx
        address_to.append(col["Addressed To"])
        cc_to.append(col["CC'd"])
    
    print(str(file_names))
    

    message = outlook.CreateItem(0) # 0 is the code for a mail item 
    message.Display()

    message.To = convert(address_to)
    message.CC = convert(cc_to)

    #This "if" statment is so we can have a different email body for Kathy. If in the future, there needs to be other unique email content to specific person, add an "if" statement formated like this.
    if (email_to_Fullname_converter(address_to) == 'kathy.ireland'):         #If Kathy is no longer in charge, change 'kathy.ireland' to the name of the new director seperated by '.'
        message.Subject = 'TCM ' + str(month_Quarter[dateTime.month-1]) + ' results'
        line1 = 'Hi ' + email_to_name_converter(address_to) + ','
        line2 = 'Hope you are doing well.'
        line3 = "Please see attached for your team's " + str(month_Quarter[dateTime.month-1]) + " results."
        line4 = 'Let me know if you or your team have any questions.'
        line5 = 'Thanks,'
        line6 = "Fatima Asif | Data Analyst | Corporate Client Group | Royal Bank of Canada | 200 Bay Street, Toronto, ON, M5V 2J2"
        line7 = "Tel: 416-302-8079| Email:fatima.asif@rbccm.com "
        
    
    else:
        message.Subject = 'Team ' + email_to_name_converter(address_to) + ' ' + str(month_Quarter[dateTime.month-1]) + ' results'
        line1 = 'Hi ' + email_to_name_converter(address_to) + ','
        line2 = 'Hope you are doing well.'
        line3 = "Please see attached for your team's " + str(month_Quarter[dateTime.month-1]) + " portfolios."
        line4 = 'Let me know if you or your team have any questions.'
        line5 = 'Thanks,' 
        line6 = "Fatima Asif | Data Analyst | Corporate Client Group | Royal Bank of Canada | 200 Bay Street, Toronto, ON, M5V 2J2"
        line7 = "Tel: 416-302-8079| Email:fatima.asif@rbccm.com "
        

    message.Body = ("%s \n \n%s \n \n%s \n \n%s \n \n \n%s \n \n%s \n%s" % (line1,line2,line3,line4,line5,line6,line7))    #if you want to add another line, name sure you add "%s" and line6 to the brackets.


    #The following loop is to include all the necessary attachments for each email, it runs through each CSV file and looks for attachments until there are none left.
    for attachment_name in file_names:   
        print(str(attachment_name))
        if (str(attachment_name)):           
            attachment_path = pathlib.Path('C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\AccountManagers\\' + str(attachment_name)) #VARIABLE: location of the attachments
            
            message.Attachments.Add(str(attachment_path))
            
    #message.save()                 #This line will save the emails to your drafts 
    #message.send()                 #This line will send all emails when running this code.

