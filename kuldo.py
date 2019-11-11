from openpyxl import load_workbook
from string import Template
import smtplib
from email.message import EmailMessage
import sys

me = 'macsaba97@gmail.com'
fileName = 'Matematika G1 információk.txt'
dataFileName = 'file.xlsx'
name = 'Marosi Csaba'
user = 'macsaba97'

inp = input('Do you want to send the mails? [y] / [n]: ')

if inp != 'y':
   sys.exit()
#A szöveg pattern beolvasása
input()
textfile = open(fileName, 'r')
textPattern = textfile.read()
subj = textfile.name.split('.', 1)[0]
textfile.close()
input()

#pw
pw = open('pass.txt', 'r').read()

#email server
server = smtplib.SMTP('smtp.gmail.com', 587)

#Next, log in to the server
server.connect('smtp.gmail.com', 587)
server.ehlo()
server.starttls()
server.ehlo()
server.login(user, pw)

# Create a text/plain message

#XLS beolvasás
wb = load_workbook(filename=dataFileName, read_only=True)
ws = wb['Munka1']

#Üzenetek összeállítása:
for row in ws.rows:
   #szöveg pattern lemásolása
   text = textPattern
   msg = EmailMessage()

   #címzett
   if(row[0].value == None):
      break;
   rec = row[0].value

   #behelyettesítés
   for i in range(1,len(row)):
      if row[i].value == None:
         break;
      text = text.replace('$' + str(i), str(row[i].value))
   msg.set_content(text)
   msg['To'] = rec
   msg['Subject'] = subj
   msg['From'] = name + '<' + me + '>' 
   try:
      server.send_message(msg)
      print('Mail sent to: ' + rec)
   except:
      print('Mail could not be sent to: ' + rec)
server.quit()  

input()
