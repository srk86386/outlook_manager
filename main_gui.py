# -*- coding: utf-8 -*-
"""
Created on Thur Jun 14 12:48:25 2020

@author: rahulkumasingh
"""


from tkinter import Tk, Text, TOP, BOTTOM, END, BOTH, CENTER, X, N, LEFT,RIGHT, RAISED, GROOVE, StringVar,PhotoImage
from tkinter.ttk import Frame, Label, Entry, Button, Style, Radiobutton
#from tkinter import*

import configparser

config = configparser.ConfigParser()
config.read('outlook_reader.config')

config_senders = (", ").join([i.strip() for i in config['COMMON']['senders'].split(",")])
config_recipients = (", ").join([i.strip() for i in config['COMMON']['recipients'].split(",")])
save_attach = int(str(config['COMMON']['save_attachments']).strip())
subject_fltr = str(config['COMMON']['subject_contains']).strip()
last_subject = str(config['SYSTEM']['last_message_subject']).strip()
last_mail_time = str(config['SYSTEM']['last_message_time']).strip()



class Example(Frame):

    def __init__(self):
        super().__init__()

        self.initUI()

    def update_config(self, section, key, value):
        #config['SYSTEM']['last_message_subject'] = "mail subject" #example
        config[section][key] = value
        
        # Writing our configuration file to 'example.ini'
        with open('outlook_reader.config', 'w') as configfile:
                config.write(configfile)
        pass
    
    def initUI(self):
        left_label_size=20
        right_item_size=80
        
        v=StringVar()
        def insert_into_texts():
            entry_recipients.insert(END, f"{config_recipients}")
            entry_senders.insert(END, f"{config_senders}")
            entry_subjesct.insert(END, f"{subject_fltr}")
            if save_attach == 1:
                r1.invoke()
            else:
                r2.invoke()
                
        def sel():
           s=v.get()
           if s=="1":
               lbl_status.config(text="Download attachment is enabled.")
               self.update_config('COMMON','save_attachments','1')
           else:
               lbl_status.config(text="Download attachment is disabled.")
               self.update_config('COMMON','save_attachments','0')
               
        def update_values():
            update_string ="Updated"
            senders = entry_senders.get()
            if senders == config_senders:
                pass
            else:
                self.update_config('COMMON','senders',senders)
                update_string = update_string + " senders list"
                
            
            recievers = entry_recipients.get()
            if recievers == config_recipients:
                pass
            else:
                self.update_config('COMMON','recipients',recievers)
                update_string = update_string + " recipients list"

            subject_contains = entry_subjesct.get()
            if subject_contains == subject_fltr:
                pass
            else:
                self.update_config('COMMON','subject_contains',subject_contains)
                update_string = update_string + " subject_fltrs"
            update_string = update_string+"."
            lbl_status.config(text=update_string)
        
        self.master.title("MailBox Manager")
        
        self.style = Style()
        self.style.theme_use("default")
        Style().configure("TButton",font='serif 10')
        Style().configure("TLabel",font='serif 10')
        Style().configure("TEntry",font='serif 10')
        
        self.pack(fill=BOTH, expand=True)

        
        frame0 = Frame(self)
        frame0.pack(fill=X)
        
        lbl_recipients = Label(frame0, font = "serif 20 bold italic underline", anchor=CENTER, text="CONFIGURATIONS", width=20)
        lbl_recipients.pack(side=TOP)
        logo = PhotoImage(file="res/configuration.png")
        w1 = Label(frame0, image=logo).pack(side=RIGHT)
        
        
        frame11 = Frame(self)
        frame11.pack(fill=X)
        
        lbl_senders = Label(frame11, text="Senders", width=left_label_size)
        lbl_senders.pack(side=LEFT, padx=10, pady=10)

        entry_senders = Entry(frame11)
        entry_senders.pack(fill=X, padx=5, expand=True)
        
        frame1 = Frame(self)
        frame1.pack(fill=X)
        
        lbl_recipients = Label(frame1, text="Recipients", width=left_label_size)
        lbl_recipients.pack(side=LEFT, padx=10, pady=10)

        entry_recipients = Entry(frame1)
        entry_recipients.pack(fill=X, padx=5, expand=True)

        frame2 = Frame(self)
        frame2.pack(fill=X)

        lbl_subject = Label(frame2, text="Subject Contains", width=left_label_size)
        lbl_subject.pack(side=LEFT, padx=10, pady=10)

        entry_subjesct = Entry(frame2)
        entry_subjesct.pack(fill=X, padx=5, expand=True)


        frame3 = Frame(self)
        frame3.pack(fill=X)

        lbl_attachment = Label(frame3, text="Download attachments", width=left_label_size)
        lbl_attachment.pack(side=LEFT, anchor=N, padx=10, pady=10)
        
        r1=Radiobutton(frame3,text="True",variable=v,value="1",command=sel)
        #r1.pack(anchor=W)
        r1.pack(fill=X, padx=10, pady=10, side=LEFT, expand=True)
        r2=Radiobutton(frame3,text="False",variable=v,value="0",command=sel)
        #r2.pack(anchor=W)
        r2.pack(fill=X, padx=10, pady=10, side=LEFT, expand=True)
        
        #lbl3 = Label(frame3, text="Review", width=6)
        #lbl3.pack(side=LEFT, anchor=N, padx=5, pady=5)

        #txt = Text(frame3)
        #txt.pack(fill=BOTH, pady=5, padx=5, expand=True)

        frame12 = Frame(self)
        frame12.pack(fill=BOTH)
        lbl_last_message_time = Label(frame12, text="Last mail date time", width=left_label_size)
        lbl_last_message_time.pack(side=LEFT, anchor=N, padx=10, pady=10)
        lbl_txt_last_message_time = Label(frame12, text=last_subject, foreground="blue", borderwidth=4, background="#C0C0C0")
        lbl_txt_last_message_time.pack(fill=X, side=LEFT, anchor=N,padx=10, pady=10)

        frame13 = Frame(self)
        frame13.pack(fill=BOTH)
        lbl_last_message_subject  = Label(frame13, text="Last mail subject", width=left_label_size)
        lbl_last_message_subject.pack(side=LEFT, anchor=N, padx=10, pady=10)
        lbl_txt_last_message_subject = Label(frame13, text=last_mail_time, foreground="blue", borderwidth=4, background="#C0C0C0")
        lbl_txt_last_message_subject.pack(fill=X, side=LEFT, anchor=N,padx=10, pady=10)

        
        frame4 = Frame(self, relief=GROOVE, borderwidth=1)
        frame4.pack(fill=BOTH)
        
        okButton = Button(frame4, text="Rules")
        okButton.pack(padx=10, pady=10, side=LEFT)
    
        closeButton = Button(frame4, text="Close",command=self.destroy)
        closeButton.pack(side=RIGHT, padx=10, pady=10)
        okButton = Button(frame4, text="Update", command=update_values)
        okButton.pack(side=RIGHT, padx=10, pady=10)

        frame5 = Frame(self)
        frame5.pack(fill=BOTH, expand=True)
        lbl_status = Label(frame5, width=100)
        lbl_status.pack(side=LEFT, padx=10, pady=10)
        
        insert_into_texts()

    
def main():

    root = Tk()
    root.geometry("800x420+200+200")
    app = Example()
    root.mainloop()


if __name__ == '__main__':
    main()
