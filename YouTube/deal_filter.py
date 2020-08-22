"""
    Script that moves all files in a folder that contain the word 'deal'
    to a 'review' folder
"""

import win32com.client as client

outlook = client.Dispatch('Outlook.Application')

namespace = outlook.GetNameSpace('MAPI')

folder = namespace.PickFolder()

folderpath = folder.FolderPath
account = folderpath[2:].split('\\')[0]
account_folder = namespace.Folders[account]

if 'JunkStuff' not in account_folder.Folders:
    junk = account_folder.Folders.Add('JunkStuff')
else:
    junk = account_folder.Folders['JunkStuff']

to_move = [mail_item for mail_item in folder.Items
 if 'deal' in mail_item.Body.lower()]

if to_move:
    for item in to_move:
        item.Move(junk)