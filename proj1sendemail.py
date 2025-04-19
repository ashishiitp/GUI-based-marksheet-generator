import smtplib
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
def sendemail(): 
     i=-1
     f= open('./sample_input/responses.csv','r')
     with f:
       reader1=csv.reader(f)
       for row1 in reader1:
                i=i+1  
                if i==0: 
                  continue
                mail_content = '''

        '''
        #The mail addresses and password
                sender_address = 'none@gmail.com'
                sender_pass = ''
                receiver_address1 = row1[1]
                receiver_address2 = row1[4]
                #Setup the MIME
                message = MIMEMultipart()
                message['From'] = sender_address
                message['To'] = receiver_address1
                message['Cc'] = receiver_address2
                message['Subject'] = 'quiz marks'
                #The subject line
                #The body and the attachments for the mail
                message.attach(MIMEText(mail_content, 'plain'))
                attach_file_name = './my output/'+row1[6]+'.xlsx'
                attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
                payload = MIMEBase('application', 'octate-stream')
                payload.set_payload((attach_file).read())
                encoders.encode_base64(payload) #encode the attachment
                #add payload header with filename
                payload.add_header("Content-Disposition", f"attachment; filename="+row1[6]+'.xlsx')
                message.attach(payload)
                #Create SMTP session for sending the mail
                session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
                session.starttls() #enable security
                session.login(sender_address, sender_pass) #login with mail_id and password
                text = message.as_string()
                session.sendmail(sender_address, receiver_address1, text)
                session.sendmail(sender_address, receiver_address2, text)
                session.quit()
                # print('Mail Sent')