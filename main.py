import win32com.client

#initialize the outlook object
outlook = win32com.client.Dispatch("Outlook.Application")

path = r'C:\AUTOMATION-PROJECTS\envio_email_from_a_template\template.msg' #enter your .msg file path

#Create a new e-mail from your .msg file
msg = outlook.CreateItemFromTemplate(path)

#Now you can make changes to the Subject, Body, and so on...

#I will just display the e-mail in my screen...
msg.Display()

#Linkedin
#https://br.linkedin.com/in/jo%C3%A3ovitormartinstst