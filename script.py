import pandas as pd
import smtplib, ssl, imaplib, time

from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from ms import html1
from config import config

#takes mailList as list of mails, beginningNumber as start point for adding mail address to list and totalMailAddressNum
# as how many mail addresses should bccList has

def add_To_Bcc(mailList, beginningNumber, totalMailAddressNum):
    bccList=[]
    if beginningNumber > len(mailList):
        return bccList
    lastNum = min(beginningNumber + totalMailAddressNum -1 , len(mailList))

    for i in range(beginningNumber,lastNum):
        bccList.append(formataddr(("",mailList[i])))
        
    return bccList

def add_To_Sent_Box(message,sentAddressList):
    message["Bcc"] = ','.join(sentAddressList)
    text = message.as_string()
    imap = imaplib.IMAP4_SSL(config.smtpServer, 993) #993, 995
    imap.login(config.senderAddress, config.password)
    imap.append('Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), text.encode('utf8'))
    imap.logout()

def prep_Mail():
    message = MIMEMultipart("alternative")
    message["Subject"] = config.subject
    message["From"] = formataddr((config.senderName,config.senderAddress))
    msgBodyPart = MIMEText(html1, "html")
    message.attach(msgBodyPart)
    message["To"] = formataddr((config.sentToName, config.sentToAddress))

    return message

#Email List
e = pd.read_excel("Email.xlsx")
emails = e['Emails'].values

server = smtplib.SMTP(config.smtpServer,config.portNum)
server.starttls()
server.login(config.senderAddress, config.password)

count = 0
while count <= len(emails):
#Make ready the message body
    message = prep_Mail()
#Make ready the Bcc email list
    bcc = add_To_Bcc(emails,count,config.bccMailAddressNum)
    count += config.bccMailAddressNum
#Make ready the all sent email list and sent the mail to them
    sentAddressList= [config.sentToAddress] + bcc
    server.sendmail(config.senderAddress, sentAddressList, message.as_string())
# Add sent mail to sent box part
    add_To_Sent_Box(message, sentAddressList[1:])

server.quit()
print("Done!")