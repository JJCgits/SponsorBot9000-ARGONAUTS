
# Python code to send email to a list of
# emails from a spreadsheet
 
# import the required libraries
import pandas as pd
import email, smtplib, ssl
import os
import mimetypes
from email.message import EmailMessage

excel_data_df = pd.read_excel('my_workbook.xls', sheet_name=0)

companyEmails = excel_data_df['Email'].tolist()
companyNames = excel_data_df['Company'].tolist()
newCompanyNames = []
newCompanyEmails = []
for name in companyNames:
    if(str(name) != "nan"):
        newCompanyNames.append(name)
for companyEmail in companyEmails:
    if(str(companyEmail) != "nan"):
        newCompanyEmails.append(companyEmail)
# change these as per use
phoneNumber = "(947) 226-0436"
studentEmail = "jjchiyezhan26@gmail.com"
my_email = "jjchiyezhan26@gmail.com"
my_password = "ppuecssikikfjwii"
subject="Troy Robotics Team Partnership"
name = "Joshua Chiyezhan"
print(newCompanyEmails)
print(newCompanyNames)

for i in range(len(newCompanyEmails)):
    companyNameX = newCompanyNames[i]
    companyEmailX = newCompanyEmails[i]

    my_message = str('''Hello, ''' + str(companyNameX) + "! \n\n" + '''My name is ''' + str(name) + ''' and I'm with the Troy Argonauts, a FIRST Robotics team out of the Troy School District. Our team competes in the FIRST Robotics Competition. We get a new challenge each year for which we design, build, and program a 120 lb. robot to play the game. We compete at state, national, and world level events and we learn skills that prepare us for careers in STEM fields. Our team is looking for partners to help cover our registration fees and robot build expenses. \n\nIf your company, ''' + str(companyNameX) + ''',  is interested in supporting K12 students in their STEM pursuits, please let us know. My phone number is ''' + phoneNumber + ''', and my head coach, Srini Simhan’s phone number is (248) 930-5827. In addition, my email is ''' + str(my_email) + ''', and the team’s email address is troyargonauts@gmail.com. \n\nBest Regards,\nJoshua Chiyezhan''')
    ccEmail = "troyargonautsgmail.com"
    attach = ("sponsorPacket.pdf")

    filenamePath = "sponsorPacket.pdf"  # In same directory as script
    mime_type, _ = mimetypes.guess_type(filenamePath)
    mime_type, mime_subtype = mime_type.split('/', 1)
    em = EmailMessage()
    em['From'] = my_email
    em['To'] = companyEmailX
    em['Subject'] = subject
    em['Cc'] = ccEmail
    em.set_content(my_message)
    context = ssl.create_default_context()
    with open(filenamePath, 'rb') as ap:
        em.add_attachment(ap.read(), maintype=mime_type, subtype=mime_subtype,
                                filename=os.path.basename(filenamePath))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login("jjchiyezhan26@gmail.com", my_password)
        # TODO: Send email here
        server.sendmail(my_email, companyEmailX, em.as_string())

