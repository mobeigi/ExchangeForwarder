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

# Set CWD to script directory
os.chdir(sys.path[0])

# Setup config
config = configparser.ConfigParser()
config.read('config.ini')
send_mode = config['DEFAULT']['SEND_MODE']

# Authenticate
credentials = Credentials(config['DEFAULT']['USERNAME'], config['DEFAULT']['PASSWORD'])
configuration = Configuration(server=config['DEFAULT']['SERVER'], credentials=credentials)

account = Account(primary_smtp_address=config['DEFAULT']['PRIMARY_SMTP_ADDRESS'], config=configuration, autodiscover=False, access_type=DELEGATE)
to_email = config['DEFAULT']['TO_EMAIL']

# Connect to smtp server
if send_mode == "smtp":
    try:
        smtp_con = smtplib.SMTP_SSL(config['DEFAULT']['SMTP_HOST'], config['DEFAULT']['SMTP_PORT'])
        smtp_con.ehlo()
        smtp_con.login(config['DEFAULT']['SMTP_SENDER_EMAIL'], config['DEFAULT']['SMTP_SENDER_PASSWORD'])
    except:
        exit('Exception occured')

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
    text = MIMEText(item.text_body, 'plain')
    html = MIMEText(item.unique_body, 'html')
    msg.attach(text)
    msg.attach(html)
    
    # Attach all attachments
    for attachment in item.attachments:
        if isinstance(attachment, FileAttachment):
            # Determine if internal image or external file attachment
            if attachment.content_type is not None and attachment.content_type.split('/')[0] == 'image':
                img = MIMEImage(attachment.content, _subtype=attachment.content_type.split('/')[1])
                img.add_header('Content-ID', '<' + attachment.content_id + '>')
                
                if attachment.is_inline:
                    img.add_header('Content-Disposition', f'inline; filename="{attachment.name}"')
                
                msg.attach(img)
            else:
                file = MIMEApplication(attachment.content)
                file.add_header('Content-Disposition', f'attachment; filename="{attachment.name}"')
                
                msg.attach(file)
       
    try:
        # Delete MeetingRequest
        if isinstance(item, MeetingRequest):
            item.move_to_trash()
        # Mark emails as read
        else:
            item.is_read = True
            item.save()
    except:
        print("Error marking email as processed")
        continue
    
    # Send email
    print("About to send email to: " + to_email + " from : " + msg["From"])
    
    if send_mode == "sendmail":
        p = Popen(["/usr/sbin/sendmail", "-t", "-oi", to_email], stdin=PIPE)
        p.communicate(msg.as_bytes())
    elif send_mode == "smtp":
        smtp_con.sendmail(config['DEFAULT']['FROM_EMAIL'], to_email, msg.as_string())
    
    print("Successfully sent email to: " + to_email + " from : " + msg["From"])
    
    # Sleep to prevent flooding/rate limiting
    time.sleep(1)
    
# Close connection
if send_mode == "smtp":
    smtp_con.close()
