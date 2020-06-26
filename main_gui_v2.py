# -*- coding: utf-8 -*-
"""
Created on Thur Jun 14 12:48:25 2020

@author: rahulkumasingh
"""


from tkinter import Tk, Text, TOP, BOTTOM, END, BOTH, CENTER, X, N, LEFT,RIGHT, RAISED, GROOVE, StringVar,PhotoImage
from tkinter.ttk import Frame, Label, Entry, Button, Style, Radiobutton
from tkinter import font  as tkfont
from tkinter.messagebox import showinfo, showerror
#from tkinter import*

import configparser
import db_manager

config = configparser.ConfigParser()
config.read('outlook_reader.config')

config_senders = (", ").join([i.strip() for i in config['COMMON']['senders'].split(",")])
config_recipients = (", ").join([i.strip() for i in config['COMMON']['recipients'].split(",")])
save_attach = int(str(config['COMMON']['save_attachments']).strip())
subject_fltr = str(config['COMMON']['subject_contains']).strip()
last_subject = str(config['SYSTEM']['last_message_subject']).strip()
last_mail_time = str(config['SYSTEM']['last_message_time']).strip()

dbmngr = db_manager.DB_Manager()

class Outlook_manager(Tk):
    dbmngr = None
    def __init__(self, *args, **kwargs):
        Tk.__init__(self, *args, **kwargs)
        # connect with db        
        self.title_font = tkfont.Font(family='Helvetica', size=18, weight="bold", slant="italic")

        # the container is where we'll stack a bunch of frames
        # on top of each other, then the one we want visible
        # will be raised above the others
        container = Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (Config_window, Rules_window,Create_Rules_window):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame

            # put all of the pages in the same location;
            # the one on the top of the stacking order
            # will be the one that is visible.
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("Config_window")

    def show_frame(self, page_name):
        '''Show a frame for the given page name'''
        frame = self.frames[page_name]
        frame.tkraise()

    
    
class Config_window(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.controller = controller
        self.initUI(self.controller)

    def update_config(self, section, key, value):
        #config['SYSTEM']['last_message_subject'] = "mail subject" #example
        config[section][key] = value
        
        # Writing our configuration file to 'example.ini'
        with open('outlook_reader.config', 'w') as configfile:
                config.write(configfile)
        pass
    
    def initUI(self, controller):
        left_label_size=20
        right_item_size=80
        
        redio_value=StringVar()
        def insert_into_texts():
            entry_recipients.insert(END, f"{config_recipients}")
            entry_senders.insert(END, f"{config_senders}")
            entry_subjesct.insert(END, f"{subject_fltr}")
            if save_attach == 1:
                r1.invoke()
            else:
                r2.invoke()
                
        def selection_control():
           s=redio_value.get()
           print(f"value_from_file- {save_attach},value - {s}, type - {type(s)}")
           if s=="1":
               self.update_config('COMMON','save_attachments','1')
               lbl_status.config(text="Download attachment is enabled.")
               print(f"{s}Download attachment is enabled.")
           elif s=="0":
               self.update_config('COMMON','save_attachments','0')
               lbl_status.config(text="Download attachment is disabled.")
               print("Download attachment is disabled.")
               
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
            selection_control()
        
        #self.master.title("MailBox Manager")
        
        #self.style = Style()
        #self.style.theme_use("default")
        #Style().configure("TButton",font='serif 10')
        #Style().configure("TLabel",font='serif 10')
        #Style().configure("TEntry",font='serif 10')
        
        #self.pack(fill=BOTH, expand=True)

        
        frame0 = Frame(self)
        frame0.pack(fill=X)
        
        lbl_recipients = Label(frame0, font = "serif 20 bold italic underline", anchor=CENTER, foreground="#4278F6", text="CONFIGURATIONS", width=20)
        lbl_recipients.pack(side=TOP)
        #logo = PhotoImage(file="res/configuration.png")
        #w1 = Label(frame0, image=logo).pack(side=RIGHT)
        
        
        frame11 = Frame(self)
        frame11.pack(fill=X)
        
        lbl_senders = Label(frame11, text="Senders", width=left_label_size, font=('Arial',12))
        lbl_senders.pack(side=LEFT, padx=10, pady=10)

        entry_senders = Entry(frame11,font=('Arial',12))
        entry_senders.pack(fill=X, padx=5, expand=True)
        
        frame1 = Frame(self)
        frame1.pack(fill=X)
        
        lbl_recipients = Label(frame1, text="Recipients", font=('Arial',12), width=left_label_size)
        lbl_recipients.pack(side=LEFT, padx=10, pady=10)

        entry_recipients = Entry(frame1, font=('Arial',12))
        entry_recipients.pack(fill=X, padx=5, expand=True)

        frame2 = Frame(self)
        frame2.pack(fill=X)

        lbl_subject = Label(frame2, text="Subject Contains", font=('Arial',12), width=left_label_size)
        lbl_subject.pack(side=LEFT, padx=10, pady=10)

        entry_subjesct = Entry(frame2)
        entry_subjesct.pack(fill=X, padx=5, expand=True)


        frame3 = Frame(self)
        frame3.pack(fill=X)

        lbl_attachment = Label(frame3, text="Download attachments", font=('Arial',12), width=left_label_size)
        lbl_attachment.pack(side=LEFT, anchor=N, padx=10, pady=10)
        
        r1=Radiobutton(frame3,text="True",variable=redio_value,value="1",command=selection_control)
        #r1.pack(anchor=W)
        r1.pack(fill=X, padx=10, pady=10, side=LEFT, expand=True)
        r2=Radiobutton(frame3,text="False",variable=redio_value,value="0",command=selection_control)
        #r2.pack(anchor=W)
        r2.pack(fill=X, padx=10, pady=10, side=LEFT, expand=True)
        
        #lbl3 = Label(frame3, text="Review", width=6)
        #lbl3.pack(side=LEFT, anchor=N, padx=5, pady=5)

        #txt = Text(frame3)
        #txt.pack(fill=BOTH, pady=5, padx=5, expand=True)

        frame12 = Frame(self)
        frame12.pack(fill=BOTH)
        lbl_last_message_time = Label(frame12, text="Last mail date time", font=('Arial',12), width=left_label_size)
        lbl_last_message_time.pack(side=LEFT, anchor=N, padx=10, pady=10)
        lbl_txt_last_message_time = Label(frame12, text=last_subject, foreground="blue", borderwidth=4, background="#C0C0C0")
        lbl_txt_last_message_time.pack(fill=X, side=LEFT, anchor=N,padx=10, pady=10)

        frame13 = Frame(self)
        frame13.pack(fill=BOTH)
        lbl_last_message_subject  = Label(frame13, text="Last mail subject", font=('Arial',12), width=left_label_size)
        lbl_last_message_subject.pack(side=LEFT, anchor=N, padx=10, pady=10)
        lbl_txt_last_message_subject = Label(frame13, text=last_mail_time, font=('Arial',12), foreground="blue", borderwidth=4, background="#C0C0C0")
        lbl_txt_last_message_subject.pack(fill=X, side=LEFT, anchor=N,padx=10, pady=10)

        
        frame4 = Frame(self, relief=GROOVE, borderwidth=1)
        frame4.pack(fill=BOTH)
        
        okButton = Button(frame4, text="Rules",command=lambda: controller.show_frame("Rules_window"))
        okButton.pack(padx=10, pady=10, side=LEFT)
    
        closeButton = Button(frame4, text="Close",command=self.quit) #self.destroy
        closeButton.pack(side=RIGHT, padx=10, pady=10)
        okButton = Button(frame4, text="Update", command=update_values)
        okButton.pack(side=RIGHT, padx=10, pady=10)

        frame5 = Frame(self)
        frame5.pack(fill=BOTH, expand=True)
        lbl_status = Label(frame5, width=100)
        lbl_status.pack(side=LEFT, padx=10, pady=10)
        
        insert_into_texts()

class Rules_window(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.controller = controller
        self.initUI(self.controller)
        #label = Label(self, text="This is page 1", font=controller.title_font)
        #label.pack(side="top", fill="x", pady=10)
        #button = Button(self, text="Go to the start page",
        #                   command=lambda: controller.show_frame("Config_window"))
        #button.pack()
    def initUI(self, controller):

        #delete rule
        def delete_rule(rule_id):
            delete_query = f"delete from rules where rule_no={rule_id};"
            print(delete_query)
            try:
                dbmngr.run_query(delete_query)
                showinfo("Delete rule","Rule was removed from database successfully.")
                create_refresh_view()
            except:
                showerror("Delete rule","Unable to remove rule from database, please try again.")

        def update_rule(rule_id):
            print(f"update rules SET where rule_no='{rule_id}';")
            #dbmngr.run_query(f"delete * from rules where rule_no='{rule_id}';")
        
        frame0 = Frame(self, padding=10)
        frame0.pack(fill=X)
        lbl_recipients = Label(frame0, font = "serif 20 bold italic underline", anchor=CENTER, foreground="#4278F6", text="RULES", width=20)        
        lbl_recipients.pack(side=TOP)
        
        frame2 = Frame(self, padding=10,relief=GROOVE, borderwidth=1)
        frame2.pack(fill=X)
            
        def create_refresh_view():
            #frame2.destroy()
            #let us create a table------------------------------------
            # get column names
            column_names = ['Rule_no', 'From_ids', 'To_ids', 'Subject_keys', 'Body_keys', 'Route_to']
            for i in range(len(column_names)):
                if i==0:
                    lbl = Label(frame2, text=column_names[i],width=8, font = "Arial 10 bold", foreground='blue', background="#C0C0C0")
                    lbl.grid(row=0, column=i)
                else:
                    lbl = Label(frame2, text=column_names[i],width=30, font = "Arial 10 bold", foreground='blue', background="#C0C0C0")
                    lbl.grid(row=0, column=i)
            # get data from db for production
            lst = dbmngr.run_query("select * from rules;")
            # find total number of rows and 
            # columns in list
            
            total_rows = len(lst) 
            total_columns = len(lst[0])

            #let us get rule IDs
            rule_ids = [i[0] for i in lst]
            
            
            # code for creating table 
            for i in range(total_rows):
                last_j=0
                for j in range(total_columns):
                    if j==0:
                        e = Entry(frame2, width=8, foreground='blue', 
                                   font=('Arial',9,'bold')) 
                    else:
                        e = Entry(frame2, width=30, foreground='blue', 
                                   font=('Arial',9,'bold')) 
                      
                    e.grid(row=i+1, column=j) 
                    e.insert(END, lst[i][j])
                    last_j = j
                #button_del = Button(frame1, text=f"delete rule_{rule_ids[i]}", command=lambda c=i: print(rule_ids[c]))
                button_del = Button(frame2, text=f"delete", command=lambda c=i: delete_rule(rule_ids[c]))
                button_del.grid(row=i+1, column=last_j+1)
            #---------------------------------------------------------
        create_refresh_view()
        
        frame1 = Frame(self, padding=10,relief=GROOVE, borderwidth=1)
        frame1.pack(fill=X)
        button = Button(frame1, text="Main Page",command=lambda: controller.show_frame("Config_window"))
        button.pack(side=LEFT)
        button = Button(frame1, text="Create Rules",command=lambda: controller.show_frame("Create_Rules_window"))
        button.pack(side=LEFT)

        button = Button(frame1, text="Refresh Rules",command=create_refresh_view)
        button.pack(side=RIGHT)
        
class Create_Rules_window(Frame):
        def __init__(self, parent, controller):
            Frame.__init__(self, parent)
            self.controller = controller
            self.initUI(self.controller)
        def initUI(self, controller):
            left_label_size=20
            s = Style()
            s.configure('my.TButton', font=('Helvetica', 19))

    
            def validate_rules():
                from_txt = txt_from.get("1.0",END)
                to_txt = txt_to.get("1.0",END)
                sub_txt = txt_sub_keys.get("1.0",END)
                body_txt = txt_body_keys.get("1.0",END)
                route_txt = txt_route_to.get("1.0",END)
                #print(f"from {from_txt}; to {to_txt}; subject {sub_txt}; body {body_txt}; route_to {route_txt};")
                return [from_txt, to_txt, sub_txt, body_txt, route_txt]

            
            def create_rule():
                rule_components = validate_rules()
                insert_query = f"""INSERT INTO rules(from_ids,to_ids,subject_keys,body_keys, route_to)
                            VALUES ('{rule_components[0]}', '{rule_components[1]}', '{rule_components[2]}'
                            , '{rule_components[3]}' , '{rule_components[4]}')
                            """
                #print(insert_query)
                try:
                    dbmngr.run_query(insert_query)
                    showinfo("Rule Creation", "Rule created successfully")
                except:
                    showerror("Rule Creation", "Error creating rule, please try again.")

                #refresh the rules view
                #rule_window = Rules_window()
                #rule_window.create_refresh_view()
                

            def reset_fields():
                txt_from.delete('1.0', END)
                txt_to.delete('1.0', END)
                txt_sub_keys.delete('1.0', END)
                txt_body_keys.delete('1.0', END)
                txt_route_to.delete('1.0', END)
                
            
            frame0 = Frame(self)
            frame0.pack(fill=X)
            lbl_recipients = Label(frame0, font = "serif 20 bold italic underline", anchor=CENTER, foreground="#4278F6", text="CREATE RULES", width=20)        
            lbl_recipients.pack(side=TOP)            

            frame2 = Frame(self)
            frame2.pack(fill=X)
            lbl_from = Label(frame2, text="From list", font=('Arial',12), width=left_label_size)
            lbl_from.pack(side=LEFT, padx=10, pady=10)
            txt_from = Text(frame2,height=2)
            txt_from.pack(fill=X, padx=5, expand=True)

            frame3 = Frame(self)
            frame3.pack(fill=X)
            lbl_to = Label(frame3, text="To list", font=('Arial',12), width=left_label_size)
            lbl_to.pack(side=LEFT, padx=10, pady=10)
            txt_to = Text(frame3,height=2)
            txt_to.pack(fill=X, padx=5, expand=True)

            frame4 = Frame(self)
            frame4.pack(fill=X)
            lbl_sub_keys = Label(frame4, text="Subject Keys", font=('Arial',12), width=left_label_size)
            lbl_sub_keys.pack(side=LEFT, padx=10, pady=10)
            txt_sub_keys = Text(frame4,height=2)
            txt_sub_keys.pack(fill=X, padx=5, expand=True)

            frame5 = Frame(self)
            frame5.pack(fill=X)
            lbl_body_keys = Label(frame5, text="Body Keys", font=('Arial',12), width=left_label_size)
            lbl_body_keys.pack(side=LEFT, padx=10, pady=10)
            txt_body_keys = Text(frame5,height=2)
            txt_body_keys.pack(fill=X, padx=5, expand=True)

            frame6 = Frame(self)
            frame6.pack(fill=X)
            lbl_route_to = Label(frame6, text="Route to list", font=('Arial',12), width=left_label_size)
            lbl_route_to.pack(side=LEFT, padx=10, pady=10)
            txt_route_to = Text(frame6,height=2)
            txt_route_to.pack(fill=X, padx=5, expand=True)

            frame1 = Frame(self, padding=10,relief=GROOVE, borderwidth=1)
            frame1.pack(fill=X)
            button = Button(frame1, text="Main Page",style='my.TButton',command=lambda: controller.show_frame("Config_window"))
            button.pack(side=LEFT, padx=8)
            button = Button(frame1, text="Manage Rules",style='my.TButton',command=lambda: controller.show_frame("Rules_window"))
            button.pack(side=LEFT, padx=8)
            button = Button(frame1, text="Reset",style='my.TButton',command=reset_fields)
            button.pack(side=RIGHT, padx=8)
            button = Button(frame1, text="Create",style='my.TButton',command=create_rule)
            button.pack(side=RIGHT, padx=8)
            
            
def main():
    root = Tk()
    root.geometry("800x420+200+200")
    root.title('TK Wizards')
    root.withdraw()
    
    app = Outlook_manager()
    root.mainloop()


if __name__ == '__main__':
    main()
