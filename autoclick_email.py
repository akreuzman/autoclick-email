#!/usr/bin/env python3

import subprocess, os, time
import gspread
import pyautogui
import pyperclip
from oauth2client.service_account import ServiceAccountCredentials

# pyautogui.PAUSE = 0.1

OUTLOOK_PATH = '/Applications/Microsoft Outlook.app'
ATTACHMENT_PATH = '/Users/snlab/Documents/Projects/fellowship-publicity/Brochure2018.pdf'

FROM_ADDRESS_COORDS = (100,162)
SEND_COORDS = (27,95)

SUBJECT_LINE = 'Research Fellowship Opportunities in Computational Neuroscience for Seniors and Recent Graduates'

# Get info from google sheet
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/snlab/Documents/Projects/fellowship-publicity/client_secret.json', scope)
client = gspread.authorize(creds)

# Open sheet
sheet = client.open('Fellowship Publicity Test').sheet1

# Get all data
result = sheet.get_all_records()

print('Preparing to create {} emails.'.format(len(result)))

# Read in email template
with open('/Users/snlab/Documents/Projects/fellowship-publicity/SimonsEmailTemplate.txt') as f:
    text = f.read()

fullSimonsEmail = text


# Loop through for every row
count = 0
for line in result:

    # Get data
    contactEmail = line['Contact Email'].strip()
    contactName = line['Contact Name'].strip()
    schoolName = line['School'].strip()
    departmentName = line['Department'].strip()
    signature = line['Signature'].strip()

    # Create email text
    emailText = fullSimonsEmail.format(contactName,schoolName,departmentName,signature)

    # Open attachment with Outlook
    subprocess.call(['open', '-a', OUTLOOK_PATH, ATTACHMENT_PATH])
    time.sleep(0.5)

    # Keystroke CTRL + <-
    pyautogui.hotkey('ctrl','left')

    # Move mouse, click email address
    pyautogui.click(FROM_ADDRESS_COORDS)

    # UP ARROW, ENTER
    pyautogui.typewrite(['up','enter'])

    # PASTE ADDRESS
    pyperclip.copy(contactEmail)
    pyautogui.hotkey('command','v')
    # pyautogui.typewrite(contactEmail)

    # tab tab tab
    pyautogui.typewrite(['tab','tab','tab'])

    # PASTE SUBJECT LINE
    pyperclip.copy(SUBJECT_LINE)
    pyautogui.hotkey('command','v')
    # pyautogui.typewrite(SUBJECT_LINE) #slow method

    # tab tab
    pyautogui.typewrite(['tab','tab'])

    # overwrite existing
    pyautogui.hotkey('command','a')

    # PASTE EMAIL BODY
    pyperclip.copy(emailText)
    pyautogui.hotkey('command','v')
    # pyautogui.typewrite(emailText) #slow method

    # CLICK TO SEND
    # pyautogui.click(SEND_COORDS) 

    # TO-DO: Update google sheet with date sent

    # Increment count
    count +=1

# DONE
print('Created {} emails.'.format(count))

