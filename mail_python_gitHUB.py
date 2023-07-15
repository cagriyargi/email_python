import pandas as pd
import smtplib

#personal information
my_name = "Email Sender Name"
my_email = "emailSender_email@loremipsum"
my_password = 'loremipsum$sid_1'

#server connection(google default)
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(my_email,my_password)

#read excel file
data=pd.read_excel(r'C:\\Users\\Admin1\\Desktop\\mails.xlsx')

#extracted data from excel file and default email message
# ad=name, soyad=surname 
for i in range(len(data)):
    gonderilecek_mail_adresi = data.loc[i]['email']
    test_email= 'Dear '+data.loc[i]['ad']+' '+data.loc[i]['soyad']+',\nLorem ipsum\nGood bye.'
    test_email = str(test_email)
    test_email = test_email.encode(encoding = 'UTF-8', errors = 'strict')
    try:
        server.sendmail(my_email, gonderilecek_mail_adresi, test_email)
        print('Email successfully sent!')
    except Exception as e:
        print('Email could not be sent'+ ' '+str(e)+'\n')

#server close
server.close()


