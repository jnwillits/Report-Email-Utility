# !/usr/bin/env python

"""
Jeff's Report Email Utility - Version 1.0
This is a utility that is executed daily by Windows Scheduler to send specific files to a distribution. It does nothing if the files do no exist. See my GitHub repository for the latest version.   Jeffrey Neil Willits     jnwillits.com
"""

import win32com.client as win32
import sys
import os
from pathlib import Path
from datetime import date


def get_paths(path):
    if os.path.exists(path):
        attachment = Path.cwd() / path
        attachment_exists = True
    else:
        attachment = None
        attachment_exists = False
    return attachment, attachment_exists


def core_tasks(file, attachment, attachment_exists):
    if attachment_exists:
        new_mail.Attachments.Add(Source=str(attachment))
        os.remove(file)


if __name__ == '__main__':

    file1 = 'filename1.pdf'
    file2 = 'filename2.pdf'
    file3 = 'filename3.pdf'

    to_email = """ < email1@gmail.com >; < email2@gmail.com > """

    attachment1, attachment1_exists = get_paths(file1)
    attachment2, attachment2_exists = get_paths(file2)
    attachment3, attachment3_exists = get_paths(file3)

    outlook = win32.gencache.EnsureDispatch('Outlook.Application')
    new_mail = outlook.CreateItem(0)
    new_mail.Subject = f'Summary Files for {date.today():%m/%d/%y}'
    new_mail.To = to_email

    core_tasks(file1, attachment1, attachment1_exists)
    core_tasks(file2, attachment2, attachment2_exists)
    core_tasks(file3, attachment3, attachment3_exists)

    if attachment1_exists or attachment2_exists or attachment3_exists:
        new_mail.Send()

sys.exit()
