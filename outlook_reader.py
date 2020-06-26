# -*- coding: utf-8 -*-
"""
Created on Wed Apr 15 12:48:25 2020

@author: krchakravarthy
"""

import win32com.client #pip install pypiwin32 to work with windows operating sysytm
import datetime
import os

path = os.getcwd()
# To get today's date in 'day-month-year' format(01-12-2017).
dateToday=datetime.datetime.today()
FormatedDate=('{:02d}'.format(dateToday.day)+'-'+'{:02d}'.format(dateToday.month)+'-'+'{:04d}'.format(dateToday.year))
print(dateToday, FormatedDate)

# Creating an object for the outlook application.
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Creating an object to access Inbox of the outlook.
inbox=outlook.GetDefaultFolder(6) # 6 is default index to access the inbox folder. to access subfolders inside the inbox .Folders.Item("Your_Folder_Name")
print(inbox)
# Creating an object to access items inside the inbox of outlook.
messages=inbox.Items

print(type(messages), len(messages))

#msg_lst = list(messages)
#print(len(msg_lst))
#print(msg_lst[0])


def save_attachments():
    
    # To iterate through inbox emails using inbox.Items object.
    for message in messages:
        #print(message)
        #break
            # Creating an object for the message.Attachments.
        attachment = message.Attachments
            # To check which item is selected among the attacments.
            # To iterate through email items using message.Attachments object.
        for attachment in message.Attachments:
            mail_date=str(message.SentOn)[0:10]
            print(mail_date)
                # To save the perticular attachment at the desired location in your hard disk.
            attachment.SaveAsFile(os.path.join(path, str(attachment)))
            break
        
    """
    for msg in messages:
        print(str(msg.SentOn))
        date = msg.SentOn.strftime("%d-%m-%y")
        if date == FormatedDate:
            print(msg.Subject, msg.SentOn)
    """
        
save_attachments() 
 
