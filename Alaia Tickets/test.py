import imaplib
import email
from email.header import decode_header
import pandas as pd
import webbrowser
import os

Artikel = "Testartikel"



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
                    print("Subject:", subject)
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
                            userlist = body.splitlines()
                            print (userlist)

                            #get user name and phone number
                            index = userlist.index('Kundendaten')
                            name = userlist[index +1]
                            tnr = userlist[index + 3]
                            mail = "to be done"

                            #get item number
                            index = userlist.index(Artikel)
                            anzahl_items = userlist[index+1]
                            anzahl_items = anzahl_items[:1]

                            # get order number
                            ordernr = userlist[2]

                            print("Ordernr:", ordernr, "User:", name, ", T-nummer: ", tnr, "will", anzahl_items, "Ticket(s)")

                            getvoucher(ordernr, name, tnr, mail, anzahl_items)




                            print("=" * 100)
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

    for i in range(len(data)):
        value = data.at[i, 'Email Sent']
        if value != None:
            ticket = data.at[i, "Voucher"]
            sendmail(ticket, name)
            crossticket(i, ordernr, name, tnr, email)
            vouchers = vouchers +1

        if vouchers >= int(anzahl_items):
            break

def sendmail(ticket, name):
    print("Sending mail ticket", ticket, "to", name)

def crossticket(i, name, tnr, email):
    print("corssing ticket at pos", i)


read_mails()