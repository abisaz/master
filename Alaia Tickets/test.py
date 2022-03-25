import imaplib
import email
from email.header import decode_header
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template


now = datetime.now()
current_time = now.strftime("%d/%m/%Y %H:%M:%S")


Artikel = "Testartikel"
sendmails = True



def read_mails():
    # Connect and login to IMAP mail server
    username = 'acdonotreply'
    password = 'kaimakkk1'
    mail_server = 'securemail-academicsurfclub-ch.prossl.de'
    imap_server = imaplib.IMAP4_SSL(host=mail_server)
    imap_server.login(username, password)

    status, messages = imap_server.select("INBOX")
    # number of top emails to fetch
    respose_code, message_numbers_raw = imap_server.search(None, 'ALL')
    message_numbers = len(message_numbers_raw[0].split()       )

    for i in range(message_numbers, 0, -1):
        # fetch the email message by ID
        res, msg = imap_server.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                # if the email message is multipart
                if msg.is_multipart() and From == '"SumUp.com" <no-reply@sumupstore.com>':
                    #print("Subject:", subject)
                    print("From:", From)
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True)
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            body = str(body)
                            body = body.replace("  ", "")
                            body = body.replace("xa0", "")
                            body = body.replace(r"\r\n", "\n")
                            body = body.replace("\n\n", "\n")
                            body = body.replace("\n\n", "\n")
                            body = body.replace("\n\n", "\n")
                            body = body.replace("\n\n", "\n")
                            body = body.replace("\n\n", "\n")
                            body = body.replace("\n\n", "\n")
                            body = body.replace("\n\n", "\n")
                            body = body.replace("@", "at")
                            userlist = body.splitlines()
                            #print (userlist)

                            #get user name and phone number
                            index = userlist.index('Kundendaten')
                            name = userlist[index +1]
                            tnr = userlist[index + 3]
                            mail = userlist[index +2]

                            #get item number
                            index = userlist.index(Artikel)
                            anzahl_items = userlist[index+1]
                            anzahl_items = anzahl_items[:1]

                            # get order number
                            ordernr = userlist[2]

                            print(ordernr, "User:", name, ", T-nummer: ", tnr, "wants", anzahl_items, "Ticket(s)")

                            getvoucher(ordernr, name, tnr, mail, anzahl_items)

                            print("*" * 100)
                            print()
                            print()
                            print()
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()


    # close the connection and logout
    imap_server.close()
    imap_server.logout()

def getvoucher(ordernr, name, tnr, email, anzahl_items):
    data= pd.read_csv("Alaiavouchers.csv")
    #print(data)
    vouchers = 0

    for i in range(len(data)): #i is the i-th voucher
        value = data.at[i, 'Email Sent'] #reat the vaule in csv at i-th row
        if pd.isna(value):  #if the value under email at i-th row is nan, then get this voucher for customer
            print("Unused voucher found")
            ticket = data.at[i, "Voucher"]  #get voucher
            crossticket(i, ordernr, name, tnr, email, ticket) #call function which completes user info @ i-th ticket
            vouchers = vouchers +1  #increment number of vouchers to keep track that on how many were send vs how many were ordered

        if vouchers >= int(anzahl_items):
            print("No more vouchers ordered in:", ordernr)
            break

def sendmail(ticket, name, email):
    email = email.replace("at", "@")

    # set up the SMTP server  ######################
    s = smtplib.SMTP(host='securemail-academicsurfclub-ch.prossl.de', port=587)
    s.starttls()
    s.login('acdonotreply', 'kaimakkk1')

    # send message to new user
    msg = MIMEMultipart()  # create a message

    # setup the parameters of the message
    msg['From'] = 'do_not_reply@academicsurfclub.ch'
    msg['To'] = email
    msg['Subject'] = "Your Alaia voucher"

    message_template = read_template('ticketmsg.txt')
    message = message_template.substitute(PERSON_NAME=name, ticket=ticket)

    # add in the message body
    msg.attach(MIMEText(message, 'plain'))

    # send the message via the server set up earlier.
    if sendmails:
        s.send_message(msg)
        print("Sent ticket", ticket, "to:", email)
    else:
        print("A wellcome mail would have been sent to: " + email)

    del msg



def crossticket(i, ordernr,  name, tnr, email, ticket):
    print("Filling ticket with ordernr, name, tnr, email and date at pos:", i)

    # reading the csv file
    df = pd.read_csv("Alaiavouchers.csv")
    # updating the column value/data
    df.loc[i, 'Order nr.'] = ordernr
    df.loc[i, 'Full Name'] = name
    df.loc[i, 'Date'] = current_time
    df.loc[i, 'Phone number'] = tnr
    df.loc[i, 'Email Sent'] = email

    sendmail(ticket, name, email)
    df.to_csv("Alaiavouchers.csv", index=False)
    print("Added user information to ticket")
    print("-"*100)


def setupSMTPserver():
    # set up the SMTP server  ######################
    s = smtplib.SMTP(host='securemail-academicsurfclub-ch.prossl.de', port=587)
    s.starttls()
    s.login('acdonotreply', 'DSu@bC7xwG&rX5nb4yzf')
    #################################################

##read textdocuments for textblocks hof and user
def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)



read_mails()