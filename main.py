'''
    Python code to send email to a list of emails from a spreadsheet
'''
# import the required libraries
# from cgi import test
import pandas as pd
from email import message
import smtplib 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import os

load_dotenv('.env')

# your email details

SERVER = 'smtp.gmail.com'  # your smtp server
PORT  = 587    # your port number
FROM  =  os.getenv('EMAIL')    # your from email id
PASS  = os.getenv('PASSWORD') # your email id password

if not FROM:
    print("no from")
if not PASS:
    print("no PASS")

if not FROM or not PASS:
    raise ValueError("Email or password is missing!")

# Authentication part
server = smtplib.SMTP(SERVER,PORT)

server.set_debuglevel(1) # set 1 to help in debugging
server.ehlo()
server.starttls() # start TLS connection which is secure connection

server.login(FROM,PASS)


  
# reading the spreadsheet
email_list = pd.read_excel('data/data.xlsx')


print('Getting the names and the emails................')
# getting the names and the emails
names = email_list['Full name']
emails = email_list['Email Address']
print(emails)

# email composing
print('Composing Email................')
# creating a message body
msg = MIMEMultipart()

msg['Subject'] = 'Invitation: FAMILY MEETING AT 4pm'
msg['From'] = FROM

print('email body settings...............')

# iterate through the records
for i in range(len(emails)):
  
    # for every record get the name and the email addresses
    name = names[i]
    email = emails[i]
  
    # the message to be emailed
    message = 'Greetings,  \n \nA kind remeinder to join our family meeting. You can find information about this meeting below.\n \nTopic: FAMILY MEETING \nTime: Feb 01 (Tuesday), 2025 \n \nJoin Zoom Meeting \nhttps://zoom.us/j/0123456789?pwd=qwerty \n\nMeeting ID: 123 456 789 \nPasscode: 0000000  \n\nLooking forward to see you today, cheers \n\nRegards'
    msg.attach(MIMEText(message, 'plain'))

    # sending the email
    server.sendmail(FROM, [email], msg.as_string())

print()
print('Email Sent .....') 

# close the smtp server
server.close()
