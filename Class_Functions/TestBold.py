from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

makeBold = '\033[1m'
endBold = "\033[0m"
OKBLUE = '\033[94m'
makeBlue = '\u001b[34m'
print(makeBlue + makeBold + "the text to bold"+ endBold)
print("the text to bold")


# text_part = MIMEText(("%s \n \n%s \n \n%s \n \n%s \n \n \n%s \n \n%s \n%s" % (line1,line2,line3,line4,line5,line6,line7)),'plain')    #if you want to add another line, name sure you add "%s" and line6 to the brackets.
# message.attach(text_part)
# html_part = """
# <div>
#     <p style="line-height: 1;"><strong><span style="color: rgb(41, 105, 176);">Fatima Asif&nbsp;</span></strong><span style="color: rgb(41, 105, 176);">| Data Analyst | Corporate Client Group | <strong>Royal Bank of Canada</strong> | 200 Bay Street, Toronto, ON, M5V 2J2</span></p>
#     <p style="line-height: 1;"><span style="color: rgb(41, 105, 176);">Tel: 416-302-8079| Email:fatima.asif@rbccm.com</span></p>
# </div>

# """
# message.attach(html_part)   