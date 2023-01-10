import imaplib
import email
import os
import json
from datetime import datetime

#Connecting to the server

imap_host = 'imap.gmail.com'
imap_user = 'yourmail@gmail1.com' #update in config file
imap_pass = 'yourpassword'

config_path = "/Users/avimunk/PycharmProjects/config/Get_gmail/"
out_path = "/Users/avimunk/Documents/bookKeeping/temp/autoinput"
if not os.path.exists(out_path):
    print("Folder does not exist, stopping code.")
    exit()

def get_sender_list():
    sender_list = []
    filename = config_path + "sender_list.json"
    with open(filename) as json_data:
        data = json.load(json_data)
    # access the data
    for item in data:
        for i in data:
            sender_list.append(i)
        return sender_list

def get_accounts_list():
    account_list = []
    filename = config_path + "accounts.json"
    with open(filename) as json_data:
        data = json.load(json_data)
    # access the data
    for item in data:
        for i in data:
            print(i['user'],i['pass'])
            account_list.append(i)
        return account_list


def download_attcments(sender,user,password):
    sender = sender
    #sender = "HOTmobile@printernet.co.il"
    # open a connection
    imap = imaplib.IMAP4_SSL(imap_host)

    # login
    imap_user = user
    imap_pass = password
    imap.login(imap_user, imap_pass)

    #List the folders in the inbox
    imap.select('Inbox')
    result, data = imap.uid('search', None, '(SENTSINCE 01-Jul-2022 FROM "{}")'.format(sender))
    print(result,data)
    #todo set the date from "SENTSINCE" as parm

    #Download attachments
    for uid in data[0].split():
        result, data = imap.uid('fetch', uid, '(RFC822)')
        raw_email = data[0][1].decode('latin-1')
        email_message = email.message_from_string(raw_email)
        counter = 1
        # download attachment
        for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()
            try:
                if bool(fileName):
                    filePath = os.path.join(out_path, fileName)
                    if not os.path.isfile(filePath) :
                        fp = open(filePath, 'w+', encoding='latin-1', buffering=1)
                        fp.write(part.get_payload(decode=True).decode('latin-1'))
                        fp.close()
            except:
                tesxtcounter = str(counter)
                filePath = os.path.join(out_path, tesxtcounter)
                fp = open(filePath, 'w+', encoding='latin-1', buffering=1)
                fp.write(part.get_payload(decode=True).decode('latin-1'))
                fp.close()
                counter = counter +1

    #Logout from the server
    imap.logout()
def main():
    accounts = get_accounts_list()
    for i in accounts:
        print(i['user'])
        user = i['user']
        password = i['pass']
        sender_list = get_sender_list()
        #print("this is the list", sender_list)
        for i in sender_list:
            print(i)
            download_attcments(i,user,password)


main()

