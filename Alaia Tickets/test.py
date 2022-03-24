import imaplib
import email
from email.header import decode_header
import webbrowser
import os

# Connect and login to IMAP mail server
username = 'acdonotreply'
password = 'kaimakkk1'
mail_server = 'securemail-academicsurfclub-ch.prossl.de'
imap_server = imaplib.IMAP4_SSL(host=mail_server)
imap_server.login(username, password)

status, messages = imap_server.select("INBOX")
# number of top emails to fetch
N = 3
# total number of emails
messages = int(messages[0])



for i in range(messages, messages-N, -1):
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
                        #print (userlist)

                        index = userlist.index('Kundendaten')
                        name = userlist[index +1]
                        #email = str(userlist[index + 2]    )
                        tnr = userlist[index + 3]
                        print(name)
                        print(tnr)

                        print("=" * 100)
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                # get the email body
                body = msg.get_payload(decode=True).decode()


# close the connection and logout
imap_server.close()
imap_server.logout()