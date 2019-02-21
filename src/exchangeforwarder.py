#!/usr/bin/env python

#
# ExchangeForwarder
# Author: Mohammad Ghasembeigi
#

import os
import sys
import configparser
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from subprocess import Popen, PIPE

from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
    EWSDateTime, EWSTimeZone, Configuration, NTLM, GSSAPI, CalendarItem, Message, \
    Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
    HTMLBody, Build, Version, FolderCollection
from exchangelib.items import MeetingRequest

import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import json

# Set CWD to script directory
os.chdir(sys.path[0])

# Setup config
config = configparser.ConfigParser()
config.read('config.ini')
send_mode = config['SMTP']['SEND_MODE']
SCOPES = 'https://www.googleapis.com/auth/calendar'

# Authenticate
credentials = Credentials(config['DEFAULT']['USERNAME'], config['DEFAULT']['PASSWORD'])
configuration = Configuration(server=config['DEFAULT']['SERVER'], credentials=credentials)

account = Account(primary_smtp_address=config['DEFAULT']['PRIMARY_SMTP_ADDRESS'], config=configuration, autodiscover=False, access_type=DELEGATE)
to_email = config['DEFAULT']['TO_EMAIL']

# Connect to smtp server
if send_mode == "smtp":
    try:
        smtp_con = smtplib.SMTP_SSL(config['SMTP']['HOST'], config['SMTP']['PORT'])
        smtp_con.ehlo()
        smtp_con.login(config['SMTP']['SENDER_EMAIL'], config['SMTP']['SENDER_PASSWORD'])
    except Exception as e:
        exit('Exception occured: ' + str(e))
        
# Initialize Google Calendar
store = file.Storage('token.json')
creds = store.get()

# Create JSON credentials.json payload
json_data = {}
json_data['installed'] = {}
json_data['installed']['client_id'] = config['GOOGLE_CALENDAR']['CLIENT_ID']
json_data['installed']['client_secret'] = config['GOOGLE_CALENDAR']['CLIENT_SECRET']
json_data['installed']['project_id'] = config['GOOGLE_CALENDAR']['PROJECT_ID']
json_data['installed']['auth_uri'] = config['GOOGLE_CALENDAR']['AUTH_URI']
json_data['installed']['token_uri'] = config['GOOGLE_CALENDAR']['TOKEN_URI']
json_data['installed']['auth_provider_x509_cert_url'] = config['GOOGLE_CALENDAR']['AUTH_PROVIDER_X509_CERT_URL']
json_data['installed']['redirect_uris'] = config['GOOGLE_CALENDAR']['REDIRECT_URIS'].split(',')

with open('credentials.json', 'w+') as outfile:
    json.dump(json_data, outfile)

if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
    creds = tools.run_flow(flow, store)
service = build('calendar', 'v3', http=creds.authorize(Http()))

# Iterate unread emails from inbox
unread = account.inbox.filter(is_read=False)
for item in reversed(unread.order_by('-datetime_received')):

    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg["From"] = item.sender.name + "<" + item.sender.email_address + ">"
    msg['Reply-To'] = msg["From"]
    
    # To recipients may be empty, in which case make the reciever the only 'to' email recipient
    if item.to_recipients is not None:
        msg["To"] = ','.join(i.name + "<" + i.email_address + ">" for i in item.to_recipients)
    elif item.received_by is not None:
        msg["To"] = item.received_by.name + "<" + item.received_by.email_address + ">"
    else:
        print("No TO recipient was found")
        continue
    
    msg["Subject"] = item.subject
    
    # Create text/html portions
    text = MIMEText(item.text_body if item.text_body is not None else "", 'plain')
    html = MIMEText(item.unique_body if item.unique_body is not None else "", 'html')
    msg.attach(text)
    msg.attach(html)
    
    # Attach all attachments
    for attachment in item.attachments:
        if isinstance(attachment, FileAttachment):
            # Determine if internal image or external file attachment
            if attachment.content_type is not None and attachment.content_type.split('/')[0] == 'image':
                img = MIMEImage(attachment.content, _subtype=attachment.content_type.split('/')[1])
                
                if attachment.content_id is not None:
                    img.add_header('Content-ID', '<' + attachment.content_id + '>')
                
                if attachment.is_inline:
                    img.add_header('Content-Disposition', f'inline; filename="{attachment.name}"')
                
                msg.attach(img)
            else:
                file = MIMEApplication(attachment.content)
                file.add_header('Content-Disposition', f'attachment; filename="{attachment.name}"')
                
                msg.attach(file)
       
    try:
        # Mark item as read and save
        item.is_read = True
        item.save(update_fields=['is_read'])  # only save is_read field to avoid issues with MeetingRequest
    except:
        print("Error marking email as processed")
        continue
    
    # Send email
    print("About to send email to: " + to_email + ", from: " + msg["From"])
    
    if send_mode == "sendmail":
        p = Popen(["/usr/sbin/sendmail", "-t", "-oi", to_email], stdin=PIPE)
        p.communicate(msg.as_bytes())
    elif send_mode == "smtp":
        smtp_con.sendmail(config['DEFAULT']['FROM_EMAIL'], to_email, msg.as_string())
    
    print("Successfully sent email to: " + to_email + ", from: " + msg["From"])
    
    # Add to calendar if MeetingRequest
    if isinstance(item, MeetingRequest):
        event = {
          'summary': item.subject,
          'location': item.location,
          'description': item.text_body if item.text_body is not None else "",
          'start': {
            'dateTime': item.start.ewsformat(),
            'timeZone': str(item._start_timezone),
          },
          'end': {
            'dateTime': item.end.ewsformat(),
            'timeZone': str(item._end_timezone),
          },
          'attendees': msg["To"].split(','),
          'reminders': {
            'useDefault': True
          },
        }

        event = service.events().insert(calendarId=config['GOOGLE_CALENDAR']['CALENDAR_ID'], body=event).execute()
        
        print("Successfully added Calendar event for: " + item.subject)
            
    # Sleep to prevent flooding/rate limiting
    time.sleep(1)
    
# Close connection
if send_mode == "smtp":
    smtp_con.close()
