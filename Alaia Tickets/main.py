import imaplib
import email
import codecs



# Connect and login to IMAP mail server
username = 'acdonotreply'
password = 'kaimakkk1'
mail_server = 'securemail-academicsurfclub-ch.prossl.de'
imap_server = imaplib.IMAP4_SSL(host=mail_server)
imap_server.login(username, password)


# Choose the mailbox (folder) to search
# Case sensitive!
imap_server.select('INBOX')  # Default is `INBOX`

# Search for emails in the mailbox that was selected.
# First, you need to search and get the message IDs.
# Then you can fetch specific messages with the IDs.
# Search filters are explained in the RFC at:
# https://tools.ietf.org/html/rfc3501#section-6.4.4
search_criteria = 'ALL'
charset = None  # All
respose_code, message_numbers_raw = imap_server.search(charset, search_criteria)
print(f'Message numbers: {message_numbers_raw}')  # e.g. ['1 2'] A list, with string of message IDs
message_numbers = message_numbers_raw[0].split()


# Find all emails in inbox
_, message_numbers_raw = imap_server.search(None, 'ALL')
for message_number in message_numbers_raw[0].split():
    _, msg = imap_server.fetch(message_number, '(RFC822)')

    # Parse the raw email message in to a convenient object
    message = email.message_from_bytes(msg[0][1])
    messagerw = message_message_raw()
    if message["from"] == '"SumUp.com" <no-reply@sumupstore.com>':
        print("New Mail from SumUp found!")
        print(f'From: {message["from"]}')
        if message.is_multipart():


            multipart_payload = message.get_payload()

            for sub_message in multipart_payload:

                my_string = sub_message.as_string()
                index = my_string.find('Kundendaten')
                kundendaten = my_string[index + 259:(index + 700)]
                kundendaten = kundendaten.replace("  ", "")
                kundendaten = kundendaten.replace("=20", "")
                undendaten = kundendaten.replace('\r', "")
                kundendaten = kundendaten.split('\n')

                name = kundendaten[0]
                email = kundendaten[1]
                tnr = kundendaten[2]
                print(name)
                print(email)
                print(tnr)

                break

    else:
        print("Other email which are not from SumUp")

imap_server.close()
imap_server.logout()