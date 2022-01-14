# Description: this program helps you automate email sending
# to dif ppl 

#Import libraries

from http import server
from logging import exception
from unicodedata import name
import pandas as pd
import numpy as np 
import datetime
import smtplib
import ssl
from email.mime.text import MIMEText as MT 
from email.mime.multipart import MIMEMultipart as MM
import os
from dotenv import load_dotenv

load_dotenv()

# Load emails list

email_df = pd.read_excel(r'/home/frank/Documents/automate_emails/project/backend/emails_bot.xlsx')
print(email_df.head())


# Create send email function

def email_funct(subject, receiver_address,name):

    # Store email receiver and sender address
    receiver = receiver_address
    sender = os.environ.get('email_from')
    sender_pswd = os.environ.get('password')

    # Create MIMEMultipart object
    msg = MM()
    msg['Subject'] = subject

    # Create HTML

    HTML= '''
    <html>
        <body> 
            <p> Hi {name}!  <br><br>
            
                We wanted to wish you an amazing and prosperous 2022 from our panda family to yours 🎉 <br>
                I am reaching out to you because we are looking for new collabs and our 2022 brand ambassadors, and we’re hoping that you could be a part of it! <br>
                We are looking to increase our reach on social media, and throughout 2021 we realised that video content is what gets people engaged the most.<br>
                This time around we are looking for original, authentic content from our influencers specifically by creating reels, stories, or anything in video format. <br>
                All that said, would you be interested in joining us on a new collab? If so, let me know and I'll send you more info 🥰 <br><br>
                Best,<br>
                The Pink Panda team 
            </p><br>
            <a href="https://pinkpandacandy.com/" target='_blank' ><img align='left' src="https://uploads-ssl.webflow.com/61e1c31d45ba1624e233829b/61e1c33a5dcd5fa5483006e7_firma%20pandas.png" width="450px" border="0"></a>
            

        </body>
    </html>
    '''.format(name=name)

    # Creat html MIMEText object
    MTObj = MT(HTML, "html")

    # Attach the MIMEText object into the mmsg container
    msg.attach(MTObj)

    # Create a secure connection with the server and send email
    # Creat the secure socket layer (SSL) contect object
    SSL_context = ssl.create_default_context()
    # Create the secure simple mail transfer protocol (SMTP) connection
    server = smtplib.SMTP_SSL(host='smtp.gmail.com',port =465, context= SSL_context)
    # Log in the email account
    server.login(sender,sender_pswd)
    # Send email
    server.sendmail(sender,receiver,msg.as_string())

# Loop over the email list

for i in range(0,len(email_df)):
    receiver = email_df['email'][i]
    name = email_df['name'][i]
    try:
        email_funct("Hello from Pink Panda!",receiver,name)
    except Exception as e:
        print('Sending email failed due to:{}'.format(e))


