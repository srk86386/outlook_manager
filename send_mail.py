import win32com.client as win32
import win32com.client


def send_mail(to_add, subject, mail_body, attachment_path=""):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to_add
    mail.Subject = subject
    #mail.Body = 'Message body'
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
    mail.HTMLBody = mail_body


    # To attach a file to the email (optional):
    
    if len(attachment_path) > 1:
        mail.Attachments.Add(attachment_path)
        
    mail.Send()
    outlook= None # free up the resources.


if __name__=="__main__":
    print("seding mail from this script")
    to_add = 'rahulkumasingh@deloitte.com'
    subject = 'Test mail'
    body = 'Hi,<br><br>This is a test mail.<br><br><br>Thanks<br>Rahul Kumar Singh'
    attachment_path = ""
    #attachment_path  = r"C:\Users\rahulkumasingh\Documents\project files\outlook_reader\Untitled.ipynb"
    send_mail(to_add, subject, body)
