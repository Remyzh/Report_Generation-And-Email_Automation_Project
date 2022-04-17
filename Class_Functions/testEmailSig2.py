import pandas as pd
import datetime
import os
import pathlib
import win32com.client
import csv
import re
outlook = win32com.client.Dispatch('outlook.application')

# Python program to convert a list
# of character
  
def convert(s):
  
    # initialization of string to ""
    new = ""
  
    # traverse in the string 
    for x in s:
        new += x
        new += '; ' 
  
    # return string 
    return new

def email_to_name_converter(email):
    strEmail = "".join(email)

    name, _, _ = strEmail.partition(".") 

    return(name.capitalize())

def email_to_Fullname_converter(email):
    strEmail = "".join(email)
    Fullname = strEmail.split("@")[0]
    return(Fullname)


def getSheetNames_function():
    os.chdir("C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test")       #VARIABLE: NEED TO CHANGE, this is the location of the overall folder

    a = open("email_list.txt", "w")

    for path, subdirs, files in os.walk(r'C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\EmailLists'):  #VARIABLE: NEED TO CHANGE, change this to location of Email List
        for filename in files:
            a.write(str(filename))
            a.write("\n")

month_Quarter = {1:'Q1 FY22', 2:'February FY22', 3:'March FY22', 4:'Q2 FY22', 5:'May FY22', 6:'June FY22', 7:'Q3 FY22', 8:'August FY22', 9: 'September FY22', 10:'Q4 FY22', 11:'November FY22', 12:'December FY22'}

dateTime = datetime.datetime.today()
print (month_Quarter[dateTime.month])



getSheetNames_function()

with open("email_list.txt", 'r') as f:                      
    listOfEmail = [line.rstrip() for line in f] 


##Main Loop

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
        file_names.append(col["File Names"] + ".xlsx")
        address_to.append(col["Addressed To"])
        cc_to.append(col["CC'd"])
    
    #print(str(file_names))
    

    message = outlook.CreateItem(0) # 0 is the code for a mail item 
    message.Display()
    message.GetInspector
    bodystart = re.search("<body.*?>", message.HTMLBody)

    message.To = convert(address_to)
    message.CC = convert(cc_to)

    if (email_to_Fullname_converter(address_to) == 'kathy.ireland'):
        message.Subject = 'TCM ' + str(month_Quarter[dateTime.month]) + ' results'
        line1 = 'Hi ' + email_to_name_converter(address_to)
        line2 = "Please see attached for your team's " + str(month_Quarter[dateTime.month]) + " results."
        line3 = 'Let me know if you or your team have any questions.'
        line4 = 'Thanks,'
        line5 = bodystart
    
    else:
        message.Subject = 'Monthly ' + str(month_Quarter[dateTime.month]) + ' results'
        line1 = 'Hi ' + email_to_name_converter(address_to) + ','
        line2 = 'Hope you are doing well.'
        line3 = "Please see attached for your team's " + str(month_Quarter[dateTime.month]) + " portfolios."
        line4 = 'Let me know if you or your team have any questions.'
        line5 = 'Thanks,' 


    
    message.Body = ("%s \n \n%s \n \n%s \n \n%s \n \n \n%s \n \n" % (line1,line2,line3,line4,line5))
   
    for attachment_name in file_names:
        attachment_path = pathlib.Path('C:\\Users\\remyz\\Documents\\RBC\\TCM Automation\\Python Test\\AccountManagers\\' + str(attachment_name)) #VARIABLE: location of the attachments
        
        print(str(attachment_path))

        message.Attachments.Add(str(attachment_path))
    
    #message.save()

