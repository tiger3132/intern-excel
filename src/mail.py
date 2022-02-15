#Import module to send email
import smtplib
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders

#Importing operating system
from subprocess import run

#Import date
from datetime import datetime

#Function
"""
Description: Send email to nattaporn at 18:30 with the excel file attached
"""
def send_attachment():

    # instance of MIMEMultipart 
    msg = MIMEMultipart() 
      
    # storing the senders email address   
    msg['From'] = ""
      
    # storing the receivers email address  
    msg['To'] = ""

    # storing the subject  
    msg['Subject'] = "Automated Daily Operation Sheet"
      
    # string to store the body of the mail 
    body = "Excel sheet in attachment"
      
    # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'plain')) 
      
    # open the file to be sent  
    filename = "testing.xlsx"
    attachment = open("/opt/PplaProject/clean.xlsx", "rb") 
      
    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 
      
    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 
      
    # encode into base64 
    encoders.encode_base64(p) 
       
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
      
    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 
    text = msg.as_string()
      
    # creates SMTP session 
    s = smtplib.SMTP('127.0.0.1', 25) 

    # sending the mail (to,from, content)
    s.sendmail("", "", text) 
      
    # terminating the session 
    s.quit()

    print('Mail Sent')
    
"""
Description: Send email containing the error created to phutana
Param: Error text
""" 
def send_error(text):

    # instance of MIMEMultipart 
    msg = MIMEMultipart() 
      
    # storing the senders email address   
    msg['From'] = "" 
      
    # storing the receivers email address  
    msg['To'] = ""

    # storing the subject  
    msg['Subject'] = "Error in automated python code"
      
    # string to store the body of the mail 
    body = text
      
    # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'plain')) 
    
    text = msg.as_string()

    # creates SMTP session 
    s = smtplib.SMTP('127.0.0.1', 25) 

    # sending the mail (to, from, content)
    s.sendmail("" , "", text) 
    
