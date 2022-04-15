from splinter import Browser
import keyboard
import time
from splinter import Browser
from selenium import webdriver
import pandas
import os
from pynput.mouse import Listener, Button, Controller

##email modules
# import the smtplib module. It should be included in Python by default
import smtplib
# import necessary packages
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template

download = True
sendmails = True
mitglieder_speichern = True


dir_name = 'export.csv'
n_new_added = 0
n_formconfirmation = 0
itemnumber = 0
coordinput = True



mouse = Controller()
xycoord = []


##read textdocuments for textblocks hof and user
def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


def setupSMTPserver():
    # set up the SMTP server  ######################
    s = smtplib.SMTP(host='securemail-academicsurfclub-ch.prossl.de', port=587)
    s.starttls()
    s.login('acdonotreply', 'DSu@bC7xwG&rX5nb4yzf')
    #################################################


##sent email with textblock to user
def send_mail_to_user(email, name, lokalgruppe):

    # set up the SMTP server  ######################
    s = smtplib.SMTP(host='securemail-academicsurfclub-ch.prossl.de', port=587)
    s.starttls()
    s.login('acdonotreply', 'kaimakkk1')

    #send message to new user
    msg = MIMEMultipart()  # create a message

    # setup the parameters of the message
    msg['From'] = 'do_not_reply@academicsurfclub.ch'
    msg['To'] = email
    msg['Subject'] = "Welcome to the ASC"

    message_template = read_template('generic_welcome_mail.txt')
    message = message_template.substitute(PERSON_NAME=name, Lokalgruppe=lokalgruppe)

    # add in the message body
    msg.attach(MIMEText(message, 'plain'))

    # send the message via the server set up earlier.
    if sendmails:
        s.send_message(msg)
        print("Sent wellcome message to: " + email)
    else:
        print("A wellcome mail would have been sent to: " + email)



    del msg


###Send email to head of finance, that a new member is registred
def send_mail_to_hof(lokalgruppenkuerzel):

    # set up the SMTP server  ######################
    s = smtplib.SMTP(host='securemail-academicsurfclub-ch.prossl.de', port=587)
    s.starttls()
    s.login('acdonotreply', 'kaimakkk1')

    ##for test
    #lokalgruppenkuerzel = 'ZH'
    resort = 'finance.' ##'finance.' ##'vicepresident.'

    #send message to the involved HOF
    msg = MIMEMultipart()  # create a message
    msg_to = resort + lokalgruppenkuerzel + '@academicsurfclub.ch'

    # setup the parameters of the message
    msg['From'] = 'do_not_reply@academicsurfclub.ch'
    msg['To'] = msg_to
    msg['Subject'] = "Neues Mitglied im ASC " + lokalgruppenkuerzel

    message_template = read_template('mail_to_hof.txt')
    message = message_template.substitute(Lokalgruppenkuerzel=lokalgruppenkuerzel)

    # add in the message body
    msg.attach(MIMEText(message, 'plain'))

    # send the message via the server set up earlier.

    # send the message via the server set up earlier.
    if lokalgruppenkuerzel == "GE" or lokalgruppenkuerzel == "FR":
        print("No HOF yet ")
    elif sendmails:
        s.send_message(msg)
        print("Sent new user reminder to: " + resort + lokalgruppenkuerzel + '@academicsurfclub.ch')
    else:
        print("A new user reminder mail would have been sent to: " + resort + lokalgruppenkuerzel + '@academicsurfclub.ch')


    del msg

##sent email with textblock to wannabeuser
def send_mail_to_wannabeuser(email, name, lokalgruppe):

    # set up the SMTP server  ######################
    s = smtplib.SMTP(host='securemail-academicsurfclub-ch.prossl.de', port=587)
    s.starttls()
    s.login('acdonotreply', 'kaimakkk1')

    #send message to new user
    msg = MIMEMultipart()  # create a message

    # setup the parameters of the message
    msg['From'] = 'do_not_reply@academicsurfclub.ch'
    msg['To'] = email
    msg['Subject'] = "Your registration at the ASC, nearly there!"

    message_template = read_template('generic_welcome_mail_wannabeuser.txt')
    message = message_template.substitute(PERSON_NAME=name, Lokalgruppe=lokalgruppe)

    # add in the message body
    msg.attach(MIMEText(message, 'plain'))

    # send the message via the server set up earlier.
    if sendmails:
        s.send_message(msg)
        print("Sent formconfirmation message to: " + email)
    else:
        print("A formconfirmation mail would have been sent to: " + email)

    del msg

###Send resume email to members@academicsurfclub.ch
def send_resume_mail(n_new_members, n_wellcommed):

    # set up the SMTP server  ######################
    s = smtplib.SMTP(host='securemail-academicsurfclub-ch.prossl.de', port=587)
    s.starttls()
    s.login('acdonotreply', 'kaimakkk1')


    #send message to the involved HOF
    msg = MIMEMultipart()  # create a message
    msg_to = "vicepresident.zh@academicsurfclub.ch"
    #msg_to = 'members@academicsurfclub.ch'

    # setup the parameters of the message
    msg['From'] = 'do_not_reply@academicsurfclub.ch'
    msg['To'] = msg_to
    msg['Subject'] = "Neu registriert: " + str(n_new_members) + ". Formconfirmation: " + str(n_wellcommed)

    message_template = read_template('resume.txt')
    message = message_template.substitute(n_new_members=n_new_members, n_wellcommed=n_wellcommed)

    # add in the message body
    msg.attach(MIMEText(message, 'plain'))

    # send the message via the server set up earlier.
    if sendmails:
        s.send_message(msg)
        print("sent resume message to HoIT")
    else:
        print("A resume mail would have been sent to HoIT")
    del msg


###Send resume email to members@academicsurfclub.ch
def send_advancedsurfer_alert(level, lokalgruppe, vorname, nachname):
    if level == "Pro":
        # set up the SMTP server  ######################
        s = smtplib.SMTP(host='securemail-academicsurfclub-ch.prossl.de', port=587)
        s.starttls()
        s.login('acdonotreply', 'DSu@bC7xwG&rX5nb4yzf')


        #send message to the involved HOF
        msg = MIMEMultipart()  # create a message
        msg_to = "sports@academicsurfclub.ch"

        # setup the parameters of the message
        msg['From'] = 'do_not_reply@academicsurfclub.ch'
        msg['To'] = msg_to
        msg['Subject'] = "Expert Surfer Alert!"

        message_template = read_template('advanced_surfer_alert.txt')
        message = message_template.substitute(Lokalgruppe=lokalgruppe, Vorname=vorname, Nachname=nachname)

        # add in the message body
        msg.attach(MIMEText(message, 'plain'))

        # send the message via the server set up earlier.
        if sendmails:
            s.send_message(msg)
            print("sent expert surfer alert to Head of Sport")
        else:
            print("A expert surfer alert mail would have been sent to HoS")
        del msg






def on_click(x, y, button, pressed):
    if pressed:
        xycoord = []
        xy = [x,y]
        xycoord.append(xy)
        print("mouse pos: {0}".format(mouse.position))
        listener.stop()

def correct(values):
    values = str(values)
    values = values.replace("ь", "ü")
    # values = values.replace("ZH", "ASC Zürich")
    values = values.replace("д", "ä")
    values = values.replace("д", "ä")
    values = values.replace("ь", "ü")
    values = values.replace("ц", "ö")
    values = values.replace("и", "è")

    return values

def get_kuerzel(lokalgruppe):
    if lokalgruppe == 'ASC Zürich':
        lokalgruppenkürzel = 'ZH'
    if lokalgruppe == 'ASC Bern':
        lokalgruppenkürzel = 'BE'
    if lokalgruppe == 'ASC Basel':
        lokalgruppenkürzel = 'BS'
    if lokalgruppe == 'ASC Luzern':
        lokalgruppenkürzel = 'LU'
    if lokalgruppe == 'ASC St.Gallen':
        lokalgruppenkürzel = 'SG'
    if lokalgruppe == 'ASC Genève':
        lokalgruppenkürzel = 'GE'
    if lokalgruppe == 'ASC Fribourg':
        lokalgruppenkürzel = "FR"

    return lokalgruppenkürzel

def student_or_alumni(value):
    # enter field gruppehinzufügen
    browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(5)>div:nth-of-type(6)>table>tbody>tr>td:nth-of-type(2)>div').click()

    # wirte alle to field
    browser.find_by_css('html>body>div:nth-of-type(4)>div:nth-of-type(2)>div>div>div>div>div>div>div:nth-of-type(3)>div>table>tbody>tr>td>input').fill("Alle")
    keyboard.press_and_release('enter')
    # oke
    time.sleep(0.5)
    browser.find_by_css('html>body>div:nth-of-type(4)>div:nth-of-type(2)>div>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>table>tbody>tr:nth-of-type(3)>td:nth-of-type(2)').double_click()
    time.sleep(0.8)

    if value == "Student":
        # enter Mitgliederbeitrag
        browser.find_by_css(
            'html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(6)>div:nth-of-type(9)>div>input').fill('29')
        keyboard.press_and_release('enter')

    if value == "Alumni":
        # enter Mitgliederbeitrag
        browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(6)>div:nth-of-type(9)>div>input').fill('50')
        keyboard.press_and_release('enter')
        # add to new group Alumnis
        browser.find_by_css(
            'html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(5)>div:nth-of-type(6)>table>tbody>tr>td:nth-of-type(2)>div').click()
        # enter field
        browser.find_by_css('html>body>div:nth-of-type(4)>div:nth-of-type(2)>div>div>div>div>div>div>div:nth-of-type(3)>div>table>tbody>tr>td>input').fill("Alumni")
        keyboard.press_and_release('enter')
        # oke
        time.sleep(0.5)
        browser.find_by_css('html>body>div:nth-of-type(4)>div:nth-of-type(2)>div>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>table>tbody>tr:nth-of-type(3)>td:nth-of-type(2)').double_click()
        time.sleep(0.5)

    # choose student/alumni
    browser.find_by_css(
        'html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(4)>div:nth-of-type(6)>div:nth-of-type(5)>div>table>tbody>tr>td>input').fill(value)
    # bestätige student/alumni auswahl mit click enter enter
    browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(4)>div:nth-of-type(6)>div:nth-of-type(5)>div>table>tbody>tr>td>input').click()
    keyboard.press_and_release('enter')
    time.sleep(0.05)
    keyboard.press_and_release('enter')
    time.sleep(0.05)


def mouseinput(vorname, nachname, adresse, plz, geburtsdatum):
    coordinput = True
    if not coordinput:
        with Listener(on_click=on_click) as listener:
            listener.join()
        keyboard.write(vorname)

        with Listener(on_click=on_click) as listener:
            listener.join()
        keyboard.write(nachname)

        with Listener(on_click=on_click) as listener:
            listener.join()
        keyboard.write(adresse)

        with Listener(on_click=on_click) as listener:
            listener.join()
        keyboard.write(plz)

        with Listener(on_click=on_click) as listener:
            listener.join()
        keyboard.write(geburtsdatum)
        coordinput = True

    else:
        timse = 0
        print(xycoord)
        time.sleep(timse)
        mouse.position = (247, 367)
        mouse.click(Button.left, 1)
        time.sleep(timse)
        keyboard.write(vorname)
        time.sleep(timse)

        mouse.position = (450, 364)
        time.sleep(timse)
        mouse.click(Button.left, 1)
        time.sleep(timse)
        keyboard.write(nachname)
        time.sleep(timse)

        mouse.position = (292, 389)
        time.sleep(timse)
        mouse.click(Button.left, 1)
        time.sleep(timse)
        keyboard.write(adresse)
        time.sleep(timse)

        mouse.position = (262, 445)
        time.sleep(timse)
        mouse.click(Button.left, 1)
        time.sleep(timse)
        keyboard.write(plz)
        time.sleep(timse)

        mouse.position = (857, 533)
        time.sleep(timse)
        mouse.click(Button.left, 1)
        time.sleep(timse)
        keyboard.write(geburtsdatum)
        time.sleep(timse)


if True:
    prof = {}
    prof['browser.download.manager.showWhenStarting'] = 'false'
    prof['browser.helperApps.alwaysAsk.force'] = 'false'
    prof['browser.download.dir'] = 'D:\OneDrive\Vereine\ASC CH\Mitgliederverwaltung Programm'
    prof['browser.download.folderList'] = 2


    prof['browser.helperApps.neverAsk.saveToDisk'] = 'text/csv'
    prof['browser.download.manager.useWindow'] = 'false'
    prof['browser.helperApps.useWindow'] = 'false'
    prof['browser.helperApps.showAlertonComplete'] = 'false'
    prof['browser.download.manager.focusWhenStarting'] = 'false'
    browser = Browser('firefox',profile_preferences=prof)
    #set window position
    browser.driver.set_window_size(1280,940)
    browser.visit('https://app.clubdesk.com/clubdesk/start')

    ##login
    browser.fill('userId', 'andri.bisaz@academicsurfclub.clubdesk.com')
    browser.fill('password', 'Kaimak:c,827x')
    browser.find_by_xpath('//*[@id="submit"]').click()


if download:
    #delete old document
    try:
        os.remove(dir_name)
    except:
        None

    time.sleep(0.5)
    #dokumente
    browser.find_by_xpath('/html/body/div[1]/div/div[2]/table/tbody/tr/td[2]/div/div[4]/div').click()

    #Click on Vorstand expand menu arrow
    time.sleep(0.5)
    #######################/html/body/div[1]/div/div[3]/div/div[4]/div[2]/div/div/div[2]/div/div[2]/div[2]/div[1]/div/div[5]/div/table/tbody/tr/td/div[xxx]/div/img[2]
    browser.find_by_xpath('/html/body/div[1]/div/div[3]/div/div[4]/div[2]/div/div[1]/div[2]/div/div[2]/div[2]/div[1]/div/div[5]/div/table/tbody/tr/td/div[11]/div[1]/img[2]').click()

    #mitglieder
    #click on A_IT_Donotchange two times
    #######################/html/body/div[1]/div/div[3]/div/div[4]/div[2]/div/div/div[2]/div/div[2]/div[2]/div[1]/div/div[5]/div/table/tbody/tr/td/div[xxx]/div[2]/div[1]/div/img[3]
    browser.find_by_xpath('/html/body/div[1]/div/div[3]/div/div[4]/div[2]/div/div[1]/div[2]/div/div[2]/div[2]/div[1]/div/div[5]/div/table/tbody/tr/td/div[11]/div[2]/div[1]/div/span[2]').click()
    time.sleep(0.5)
    browser.find_by_xpath('/html/body/div[1]/div/div[3]/div/div[4]/div[2]/div/div[1]/div[2]/div/div[2]/div[2]/div[1]/div/div[5]/div/table/tbody/tr/td/div[11]/div[2]/div[1]/div/span[2]').click()
    time.sleep(0.5)
    #click on mitglieder schweizweit
    browser.find_by_xpath('/html/body/div[1]/div/div[3]/div/div[4]/div[2]/div/div/div[2]/div/div[1]/div[2]/div[1]/div/div[2]/div[2]/div/div[3]/div[1]/table/tbody[2]/tr/td[2]/div').click()

    # click on öffnen
    browser.find_by_xpath('/html/body/div[1]/div/div[3]/div/div[4]/div[2]/div/div/div[1]/div/div/table[1]/tbody/tr[2]/td[2]/div/div/div/table/tbody/tr[1]/td[2]/div/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td[2]/div').click()

    time.sleep(0.8)
    #export
    browser.find_by_xpath('/html/body/div[1]/div/div[3]/div/div[4]/div[2]/div/div[2]/div[1]/div/div/table[3]/tbody/tr[2]/td[2]/div/div/div/table/tbody/tr[1]/td[2]').click()

    #download
    browser.find_by_xpath('/html/body/div[4]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/table/tbody/tr[2]/td[2]/div/div/table/tbody/tr/td/div').click()

    time.sleep(0.8)


#read csv, only the ones who are not registred
rows = 0
whole_dict = pandas.read_csv(dir_name, encoding='windows-1251', sep=';')
len_tot = len(whole_dict.index)
print("Total list entries ", len_tot)

#add a column with the row to later check the right boxes, altough rest of dict is missing
whole_dict["Itemnumber"] = list(range(0, len_tot))

#reorder to dict
mitglieder = whole_dict.to_dict(orient="records")

for i in range(len(mitglieder)):
    if True:
        a = 0
        lokalgruppe = correct(mitglieder[i].get('Local group'))
        lokalgruppenkürzel = get_kuerzel(lokalgruppe)
        Nachname = correct(mitglieder[i].get('Last Name'))
        Vorname = correct(mitglieder[i].get('First Name'))
        Geburtsdatum = mitglieder[i].get('Date of Birth')
        Telefon = mitglieder[i].get('Mobile number')
        email = correct(mitglieder[i].get('E-Mail'))
        Adresse = correct(mitglieder[i].get('Address'))
        Plz = mitglieder[i].get('ZIP Code')
        Ort = mitglieder[i].get('City')
        Land = mitglieder[i].get('Country')
        Matrikelnummer = mitglieder[i].get('Matriculation number')
        Uni = correct(mitglieder[i].get('University'))
        Studiengang = correct(mitglieder[i].get('Degree'))
        Student_Alumni = mitglieder[i].get('Student/Alumni')
        Surflevel = mitglieder[i].get('Surflevel')
        Mitgliederbeitrag = mitglieder[i].get('Membership fee')
        Bemerkungen = mitglieder[i].get('Comments')
        Bezahlt = mitglieder[i].get('Bezahlt')
        Registriert = mitglieder[i].get('Registriert')
        Formconfirmation = mitglieder[i].get('Formconfirmation')
        Bemerkungen = correct(mitglieder[i].get('Bemerkungen'))
        Itemnumber = mitglieder[i].get('Itemnumber') + 1

        print("User number: " + str(Itemnumber) + ", Name: " + str(Vorname))


        #register a new member
        if Bezahlt == "Ja" and Registriert == "Nein":

            # Mitglieder Registerkarte
            browser.find_by_xpath('/html/body/div[1]/div/div[2]/table/tbody/tr/td[2]/div/div[2]/div').click()
            # wechsel raus aus registerkarte
            browser.find_by_xpath('/html/body/div[1]/div/div[3]/div/div[2]/div[1]/div/ul/li[1]/a[2]/em/span/span').click()
            time.sleep(0.3)

            # Neues Mitglied erstellen
            nutzername = None
            time.sleep(0.5)
            browser.find_by_xpath('/html/body/div[1]/div/div[3]/div/div[2]/div[2]/div/div/div[1]/div/div/table[1]/tbody/tr[2]/td[2]/div/div/div/table/tbody/tr[1]/td[1]').click()
            browser.find_by_xpath('/html/body/div[4]/div/div[1]/div/span').click()
            # add to new groups, click on the field
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(5)>div:nth-of-type(6)>table>tbody>tr>td:nth-of-type(2)>div').click()
            # write in the field
            browser.find_by_css('html>body>div:nth-of-type(4)>div:nth-of-type(2)>div>div>div>div>div>div>div:nth-of-type(3)>div>table>tbody>tr>td>input').fill(lokalgruppe)
            keyboard.press_and_release('enter')
            # oke
            time.sleep(0.5)
            browser.find_by_css('html>body>div:nth-of-type(4)>div:nth-of-type(2)>div>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>table>tbody>tr:nth-of-type(3)>td:nth-of-type(2)').double_click()
            time.sleep(0.5)
            # add values to field welche loalgrupe
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(7)>div:nth-of-type(3)>div>input').fill(lokalgruppe)
            #add phone number
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div:nth-of-type(6)>div>input').fill(Telefon)
            #add email
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div:nth-of-type(15)>div>input').fill(email)
            #ort
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div>div:nth-of-type(15)>div:nth-of-type(3)>div>input').fill(Ort)
            #Land
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div>div:nth-of-type(15)>div:nth-of-type(5)>div>table>tbody>tr>td>input').fill(Land)
            #Matrikelnummer
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(6)>div:nth-of-type(12)>div>input').fill(Matrikelnummer)
            #Universität
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(6)>div:nth-of-type(3)>div>input').fill(Uni)
            #Studiengang
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(6)>div:nth-of-type(6)>div>input').fill(Studiengang)
            #student or alumni? via function
            student_or_alumni(Student_Alumni)
            #surflevel
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(7)>div:nth-of-type(6)>div>input').fill(Surflevel)
            #Eintritt definieren
            browser.find_by_xpath('/html/body/div[1]/div/div[3]/div/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div/div[1]/div[4]/div[6]/div[1]/div/table/tbody/tr/td[2]/div').click()
            browser.find_by_css('html>body>div:nth-of-type(4)>div>div>div>table:nth-of-type(2)>tbody>tr>td>div>div>table>tbody>tr:nth-of-type(2)>td:nth-of-type(2)>div>div>table>tbody>tr>td>div').click()
            #Restliche mousinputs
            mouseinput(Vorname, Nachname, Adresse, Plz, Geburtsdatum)
            time.sleep(0.3)


            # benutzername erstellen
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div:nth-of-type(2)>div>div>div:nth-of-type(4)>div:nth-of-type(9)>div>div>input').click()
            # mitglied speichern
            if mitglieder_speichern:
                browser.find_by_xpath('html[1]/body[1]/div[1]/div[1]/div[3]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[2]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[2]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/img[1]').click()

            # auf Liste wechseln
            browser.find_by_xpath('/html/body/div[1]/div/div[2]/table/tbody/tr/td[2]/div/div[4]/div').click()
            #check boxes that this user is registered
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(4)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div>div>div:nth-of-type(2)>div>table>tbody:nth-of-type(2)>tr:nth-of-type(' + str(Itemnumber) + ')>td:nth-of-type(21)>div>input').double_click()
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(4)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div>div>div:nth-of-type(2)>div>table>tbody:nth-of-type(2)>tr:nth-of-type(' + str(Itemnumber) + ')>td:nth-of-type(21)>div>input').click()

            #if "bezahlt" but no formconfirmation yet, directly sent wellcome mail!
            if Formconfirmation =="Nein":
                print("Formconfirmation is not sent, due to directly wellcomemail")
                browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(4)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div>div>div:nth-of-type(2)>div>table>tbody:nth-of-type(2)>tr:nth-of-type(' + str(Itemnumber) + ')>td:nth-of-type(22)>div>input').double_click()
                browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(4)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div>div>div:nth-of-type(2)>div>table>tbody:nth-of-type(2)>tr:nth-of-type(' + str(Itemnumber) + ')>td:nth-of-type(22)>div>input').click()

            #send wellcome mail to user:
            send_mail_to_user(email, Vorname, lokalgruppe)

            #send advcanced surfer alert to Head of Sports
            send_advancedsurfer_alert(Surflevel, lokalgruppe, Vorname, Nachname)

            #add counts
            n_new_added = n_new_added + 1

        #send a received form confirmation for new wannabeemember
        if Formconfirmation == "Nein" and Bezahlt == "Nein":
            # formconfirmation email sending to wannabeuser
            send_mail_to_wannabeuser(email, Vorname, lokalgruppe)
            # auf Liste wechseln
            browser.find_by_xpath('/html/body/div[1]/div/div[2]/table/tbody/tr/td[2]/div/div[4]/div').click()
            # check box that formconfirmation email was sent:
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(4)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div>div>div:nth-of-type(2)>div>table>tbody:nth-of-type(2)>tr:nth-of-type(' + str(Itemnumber) + ')>td:nth-of-type(22)>div>input').double_click()
            browser.find_by_css('html>body>div>div>div:nth-of-type(3)>div>div:nth-of-type(4)>div:nth-of-type(2)>div>div:nth-of-type(2)>div:nth-of-type(2)>div>div>div>div>div:nth-of-type(2)>div>table>tbody:nth-of-type(2)>tr:nth-of-type(' + str(Itemnumber) + ')>td:nth-of-type(22)>div>input').click()
            #send email to head of finance that a new User is waiting
            send_mail_to_hof(lokalgruppenkürzel)
            n_formconfirmation = n_formconfirmation + 1

#Liste speichern und schliessen
browser.find_by_xpath('/ html / body / div[1] / div / div[3] / div / div[4] / div[2] / div / div[2] / div[1] / div / div / table[1] / tbody / tr[2] / td[2] / div / div / div / table / tbody / tr[1] / td[1] / div / div / table / tbody / tr[2] / td[2] / div / div / table / tbody / tr / td[2] / div').click()


#send mail to admin to inform
send_resume_mail(n_new_added, n_formconfirmation)

#close windows
try:
    window = browser.windows[0]
    #window.close()
except:
    None
