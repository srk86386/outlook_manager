#router
import send_mail

def route_my_mail(subject, body, mail_string):
    #let us define a rule to route mail
    words_in_subject = subject.split()
    words_in_subject = [w.lower() for w in words_in_subject]
    
    words_in_body = body.split()
    words_in_body = [w.lower() for w in words_in_body]
    
    subject_keys = "attendees"
    subject_intersact = [r1 for r1 in words_in_subject if r1 in subject_keys]
    body_intersact = [r1 for r1 in words_in_body if r1 in subject_keys]
    if (len(subject_intersact) > 0) and (len(body_intersact) > 0):
        send_mail.send_mail("rahulkumasingh@deloitte.com",f"test-{mail_sub}", mail_string)
        print(f"sent mail - {subject}")
    #pass
