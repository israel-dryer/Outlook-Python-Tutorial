"""
    Script that looks for all mail in my inbox with a subject that contains 
    the words 'Python' and save the recient and subject line in a csv file
"""

import csv
from datetime import datetime
import win32com.client as client

def mail_body_search(term, folder): 
    """Recursively search all folders for email containing the search term"""
    try:
        relevant_messages = [message for message in folder.Items if term in message.Body.lower()]
    except AttributeError:
        # not items in the current folder
        relevant_messages = []

    # check for subfolders (base case)
    subfolder_count = folder.Folders.Count

    # search all subfolders
    if subfolder_count > 0:
        for subfolder in folder.Folders:
            relevant_messages.extend(mail_body_search(term, subfolder))

    return relevant_messages

# extract all python messages and save select data to file
today = datetime.today().strftime('%Y%m%d')
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNameSpace('MAPI')
messages = mail_body_search('python', namespace)

with open('py_mail_' + today + '.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    for item in messages:
        sender_email = item.SenderEmailAddress
        subject = item.subject
        writer.writerow([sender_email, subject])
