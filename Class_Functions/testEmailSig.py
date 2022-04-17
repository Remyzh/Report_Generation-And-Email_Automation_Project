import pandas as pd
import datetime
import os
import pathlib
import win32com.client
outlook = win32com.client.Dispatch('outlook.application')
import csv
import re
    
message = outlook.CreateItem(0) # 0 is the code for a mail item 


# message.body = signature
# message.Display()
message.GetInspector
bodystart = re.search("<body.*?>", message.HTMLBody)
message.HTMLBody = re.sub(bodystart.group(), bodystart.group()+"Hello this is a test body",message.HTMLBody)
print(str(message.HTMLBody))
message.Display()