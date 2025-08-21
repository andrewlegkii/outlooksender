import win32com.client as win32

def send_email(to, cc, subject, body):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.CC = cc
    mail.Subject = subject
    mail.HTMLBody = body
    mail.Send()
