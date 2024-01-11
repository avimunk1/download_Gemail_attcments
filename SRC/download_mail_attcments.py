# this scripts search the inbox for attachments sent from a list of senders and download them to the local computer

import imaplib
import email
import os
import json
import loguru as logger
import re
import random
from html2text import html2text
import pdfkit
from datetime import datetime

# set the search date
original_date_str = "01-07-2023" #this is the date from which the script will search for emails
original_date = datetime.strptime(original_date_str, "%d-%m-%Y") # Convert to datetime object
imap_date_str = original_date.strftime("%d-%b-%Y") # Convert to the format expected by IMAP (dd-Mon-yyyy)

# Connecting to the server
imap_host = 'imap.gmail.com'
imap_user = 'yourmail@gmail1.com'  # update in config file
imap_pass = 'yourpassword'
config_path = "/Users/avimunk/PycharmProjects/config/Get_gmail/"
out_path = "/Users/avimunk/Documents/bookKeeping/temp/"
if not os.path.exists(out_path):
    print("Folder does not exist, stopping code.")
    exit()

def get_sender_list():
    sender_list = []
    filename = config_path + "sender_list.json"
    with open(filename) as json_data:
        data = json.load(json_data)
    # access the data
    # for item in data:
    for i in data:
        sender_list.append(i)
    return sender_list

def get_accounts_list():
    account_list = []
    filename = config_path + "accounts.json"
    with open(filename) as json_data:
        data = json.load(json_data)

    # access the data
    for i in data:
        print(i['user'], i['pass'])
        account_list.append(i)
    return account_list

def download_attcments(sender, user, password):
    sender = sender
    # open a connection
    imap = imaplib.IMAP4_SSL(imap_host)
    # login
    imap_user = user
    imap_pass = password
    imap.login(imap_user, imap_pass)
    counter = str(random.randint(0, 9999))

    # set List of folders in the inbox to search
    imap.select('Inbox')
    result, data = imap.uid('search', None, '(SENTSINCE {} FROM "{}")'.format(imap_date_str, sender))

    # find attachments
    for uid in data[0].split():
        result, data = imap.uid('fetch', uid, '(RFC822)')
        raw_email = data[0][1].decode('latin-1')
        email_message = email.message_from_string(raw_email)

        # Check for attachments first
        has_attachments = False
        for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is not None and part.get_filename() is not None:
                has_attachments = True
        # download attachment
        if has_attachments:
            print("has attachments")
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                # set file name
                sender_name = sender.split('@')
                sender_name = sender_name[1]
                fileName = part.get_filename()

                fileName = sender_name + fileName + str(random.randint(0, 9999)) + ".pdf"
                print(fileName)
                if type(fileName) == str:
                    sender_name + fileName
                try:
                    if bool(fileName):
                        filePath = os.path.join(out_path, fileName)
                        if not os.path.isfile(filePath):
                            fp = open(filePath, 'w+', encoding='latin-1', buffering=1)
                            fp.write(part.get_payload(decode=True).decode('latin-1'))
                            fp.close()
                except:
                    tesxtcounter = str(counter)
                    filePath = os.path.join(out_path, tesxtcounter + ".pdf")
                    fp = open(filePath, 'w+', encoding='latin-1', buffering=1)
                    fp.write(part.get_payload(decode=True).decode('latin-1'))
                    fp.close()
                    counter = str(random.randint(0, 9999))
        # If no attachments, process the email body
        if not has_attachments:
            fileName = str(random.randint(0, 9999)) + ".pdf"
            print("no attachments",fileName)
            body = ""
            if email_message.is_multipart():
                for part in email_message.walk():
                    ctype = part.get_content_type()
                    cdispo = str(part.get('Content-Disposition'))

                    # skip any text/plain (txt) attachments
                    if ctype == 'text/plain' and 'attachment' not in cdispo:
                        body = part.get_payload(decode=True)  # decode
                        break
                    elif ctype == 'text/html':
                        body = part.get_payload(decode=True)
                        break
            else:
                body = email_message.get_payload(decode=True)

            # Check if body is bytes and decode if necessary
            if isinstance(body, bytes):
                try:
                    body = body.decode('utf-8')
                except:
                    print("error decoding body",fileName)
                    continue

            # Save the body content to a file
            file_name = "email_body_" + sender + str(random.randint(0, 9999))
            ctype = part.get_content_type()
            if ctype == 'text/html':
                pdf_file_name = file_name + ".pdf"
                pdf_file_path = os.path.join(out_path, pdf_file_name)
                pdfkit.from_string(body, pdf_file_path)
            else:
                txt_file_name = file_name + ".txt"
                txt_file_path = os.path.join(out_path, txt_file_name)
                with open(txt_file_path, 'w', encoding='utf-8') as file:
                    file.write(body)

    # Logout from the server
    imap.logout()


def main():
    accounts = get_accounts_list()
    for i in accounts:
        print(i['user'])
        user = i['user']
        password = i['pass']
        sender_list = get_sender_list()
        # print("this is the list", sender_list)
        for i in sender_list:
            print(i)
            download_attcments(i, user, password)


main()
