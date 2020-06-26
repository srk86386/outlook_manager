# -*- coding: utf-8 -*-
"""
Created on Thur May 28 12:48:25 2020

@author: rahulkumasingh
"""

import win32com.client #pip install pypiwin32 to work with windows operating sysytm
import datetime
import os
import send_mail
import router
import configparser
import sys
config = configparser.ConfigParser()
config.read('outlook_reader.config')



path = os.getcwd()
# To get today's date in 'day-month-year' format(01-12-2017).
dateToday=datetime.datetime.today()
FormatedDate=('{:02d}'.format(dateToday.day)+'-'+'{:02d}'.format(dateToday.month))
print(dateToday, FormatedDate)

# Creating an object for the outlook application.
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Creating an object to access Inbox of the outlook.
inbox=outlook.GetDefaultFolder(6) # 6 is default index to access the inbox folder. to access subfolders inside the inbox .Folders.Item("Your_Folder_Name")
print(inbox)
# Creating an object to access items inside the inbox of outlook.
messages=inbox.Items

def init_config():
        date_string = str(config['SYSTEM']['last_message_time']).strip()
        if len(date_string) != 0:
                date_object = datetime.datetime.strptime(date_string, "%d-%m-%Y %H-%M-%S")
                if date_object.strftime("%d%m") == datetime.datetime.now().strftime("%d%m"):
                        print("last run happened today!")
                else:
                        print(f"last run did not happened today!{date_object.strftime('%d%m'), datetime.datetime.now().strftime('%d%m')}")
                        update_config('','') # removes last updated time and subject
        else:
                pass
        #sys.exit(0)
        #pass
def save_attachments(message):
        if str(config['COMMON']['save_attachments']).strip() == '1':
                # Creating an object for the message.Attachments.
                attachment = message.Attachments
                    # To check which item is selected among the attacments.
                    # To iterate through email items using message.Attachments object.
                for attachment in message.Attachments:
                    mail_date=str(message.SentOn)[0:10]
                    print(mail_date)
                        # To save the perticular attachment at the desired location in your hard disk.
                    attachment.SaveAsFile(os.path.join(path+"/attachments", str(attachment)))
                    break
        elif str(config['COMMON']['save_attachments']).strip() == '0':
                pass
        else:
                print("Please the value of save_attachments and correct it to have 0 for no download or 1 for download.")

def update_config(mail_time, mail_subject):
        config.set('SYSTEM','last_message_time', mail_time)
        config['SYSTEM']['last_message_subject'] = mail_subject
                
        # Writing our configuration file to 'example.ini'
        with open('outlook_reader.config', 'w') as configfile:
                config.write(configfile)


def verify_subject_time_recipient(time,subjet,mail_recipients):
        should_read_mail = False
        
        senders = [i.strip() for i in config['COMMON']['senders'].split(",")]
        config_recipients = [i.strip() for i in config['COMMON']['recipients'].split(",")]
        subject_fltr = str(config['COMMON']['subject_contains']).strip()
        last_subject = str(config['SYSTEM']['last_message_subject']).strip()

        recipients_list = [r1 for r1 in config_recipients if r1 in mail_recipients] #intersaction of lists

        format1 = "%d-%m-%Y %H-%M-%S"# The format
        config_time = str(config['SYSTEM']['last_message_time']).strip()
        date1=''
        if config_time == "":
                pass
        else:
                date1 = datetime.datetime.strptime(config_time, format1)
        date2 = datetime.datetime.strptime(time, format1)
        #print(config_recipients, mail_recipients)
        #print(f"total matches found {len(recipients_list)}")
        #print(len(config_recipients), len(recipients_list))
        if (len(recipients_list) > 0) or (len(config_recipients)==1):
                #print("found recipient")
                #print(config_time, date2, date1)
                if (config_time==""):
                        #print("matched time")
                        if not subject_fltr.strip() == "":
                                if subjet.find(subject_fltr) > 0:
                                        #print("matched subject")
                                        should_read_mail= True
                                else:
                                        should_read_mail= False
                        else:
                                #print("no subject")
                                should_read_mail= True
                elif (date2 > date1) and (last_subject != subjet):
                        #print("matched time")
                        if not subject_fltr.strip() == "":
                                if subjet.find(subject_fltr) > 0:
                                        print("matched subject")
                                        should_read_mail= True
                                else:
                                        should_read_mail= False
                        else:
                                #print("no subject")
                                should_read_mail= True
                else:
                        should_read_mail= False
                        
        else:
                should_read_mail= False 
        
        return should_read_mail


def get_todays_mails(all_mails):
    msg = all_mails.GetLast()
    #mail_date = msg.ReceivedTime.strftime("%d-%m-%Y")
    counter = 1
    last_mail = []
    while msg:
        mail_date, mail_sub, mail_sender,mail_sender_name, mail_recevedOn, mail_body, mail_receiver= "","","","","","",""
        mail_date = msg.SentOn.strftime("%d-%m")
        #print(mail_date, FormatedDate)
        if FormatedDate == mail_date:
                # Reading mail information
                mail_sub = msg.Subject
                try:
                        mail_sender = msg.Sender.Address
                        mail_sender_name = msg.SenderName
                except Exception as e:
                        print(e)
                #mail_send_acc = msg.SendUsingAccount
                try:
                        mail_recevedOn = msg.ReceivedTime.strftime("%d-%m-%Y %H-%M-%S")
                except:
                        mail_recevedOn = msg.SentOn.strftime("%d-%m-%Y %H-%M-%S")
                mail_body = ''
                try:
                        mail_body = msg.HTMLBody
                except:
                        mail_body = msg.Body
                mail_receiver = ''
                try:
                        mail_receiver = msg.ReceivedByName
                except:
                        try:
                                mail_receiver = msg.Recipients.Address
                        except:
                                print("unable to get the recipients")
                #mail_recipients = msg.ReceivedByName # prints recipients name
                #mail_recipients = msg.Recipients.Address
                recipients_mail = []
                for recipient in msg.Recipients:
                        try:
                                recipients_mail.append(recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress)
                        except Exception as e:
                                print(e)
                #print(f"{counter} verified mail - {mail_sub}1")
                if verify_subject_time_recipient(mail_recevedOn, mail_sub, recipients_mail):
                        #print(f" verified mail - {mail_sub}")
                        save_attachments(msg) # decide and save the attachments
                        if counter == 1:
                                last_mail.append(mail_recevedOn)
                                last_mail.append(mail_sub)
                                pass
                        print(f"{counter} - sub:{mail_sub} || sender:{mail_sender_name} || recipients:{recipients_mail}|| receiver:{mail_receiver} || received on:{mail_recevedOn}")
                        
                        body_html = "Hi,<br><br>"
                        body_html = body_html+ f"recipients: {recipients_mail}<br>received on:{mail_recevedOn}<br><br>{mail_body}<br><br>"
                        body_html = body_html+"Rahul Kumar Singh"
                        router.route_my_mail(mail_sub, mail_body, body_html)
                        #send_mail.send_mail("rahulkumasingh@deloitte.com",f"test-{mail_sub}", body_html)
                else:
                        pass
                        #print("There are no new e-mails.")
                #print(f"{counter} - sub:{mail_sub} || sender:{mail_sender_name} || send_acc:{mail_sender} || recipients:{recipients_mail}|| receiver:{mail_receiver} || received on:{mail_recevedOn}")
            
        else:
                break
        msg = all_mails.GetPrevious()
        counter+=1
    update_config(last_mail[0],last_mail[1])

init_config()

get_todays_mails(messages)
#save_attachments()
outlook = None
config = None
