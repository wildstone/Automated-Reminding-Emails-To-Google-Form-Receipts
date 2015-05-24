####################################
# This is a python script of generating reminder emails for Google Form receipts.
# It will send reminder emails to those who didn't complete the specified Google
# Form.
# The script is based on:
#     1. http://stackoverflow.com/questions/10147455/trying-to-send-email-gmail-as-mail-provider-using-python
#     2. https://github.com/burnash/gspread
###################################



import json
import gspread

from oauth2client.client import SignedJwtAssertionCredentials

import smtplib
from email.mime.text import MIMEText
import email

class GoogleFormResponseReader:
    def __init__(self, directory_to_json_file):
        json_key = json.load(open(directory_to_json_file))
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'], scope)
        self.reader = gspread.authorize(credentials)
        self.response_sheet = None
        
    def ReadByName(self, google_form_response_name):
        self.response_sheet = self.reader.open(google_form_response_name).sheet1
        
    def GetEmailsSubmitted(self, column_index_of_emails):
        return self.response_sheet.col_values(column_index_of_emails)

class GoogleFormReminder:
    def __init__(self, my_email, my_password):
        self.my_email = my_email
        self.my_password = my_password
        self.receipts = dict()
        
    def SetReceipts(self, receipts):
        # receipts is a string of email_addresses with comma as separator
        email_addresses = receipts.split(',')
        for email in email_addresses:
            self.receipts[email] = 1
        
    def UpdateReceipts(self, emails_submitted):     
        for email in emails_submitted:
            if self.receipts.has_key(email):
                del self.receipts[email]
                
    def SendReminders(self, form_name, form_url):
        server = smtplib.SMTP('smtp.gmail.com:587')
        server.ehlo()
        server.starttls()
        server.login(self.my_email, self.my_password)
        for email in self.receipts:
            msg = MIMEText('Hello\nPlease complete ' + form_name \
                           + 'and submit it via the following link as soon as possibleurl.\n\n' \
                           + form_url)
            msg['Subject'] = '[Reminder]Please complete ' + form_name
            msg['From'] = self.my_email
            msg['To'] = email
            server.sendmail(self.my_email, email, msg.as_string())
            
# Set key_file_location which is the direct or relative location of json file used for credential
# of gspread
# To get the needed json file, read http://gspread.readthedocs.org/en/latest/oauth2.html.
# Be careful about the sixth step, you need manully share your Google Form response sheet
# with json_key['client_email']
directory_to_json_file = 'directory-to-json-file'
reader = GoogleFormResponseReader(directory_to_json_file)

# Reads Google Form response spread sheet
google_form_name = 'Sample Google Form Spreadsheet' # Set the name of the specify Google Form

# Set the url of Google Form
google_form_url = 'http://sample-url'

# Set the name of Google Form Response Spreadsheet or using the default name as following
google_form_response_name = google_form_name + ' (Responses)'

# Read Google Form response
reader.ReadByName(google_form_response_name)

# Set the index of column of emails stored in Google Form response spreadsheet
column_index_of_emails = 3
emails_submitted = reader.GetEmailsSubmitted(column_index_of_emails)

# Set the gmail used as sender email and its password
my_email = 'myemail@gmail.com'
my_password = 'mypassword'
reminder = GoogleFormReminder(my_email, my_password)

# Set receipts which is a string of email addresses with comma as separator
receipts = 'example1@gmail.com,example2@gmail.com,example3@gmail.com'
reminder.SetReceipts(receipts)

reminder.UpdateReceipts(emails_submitted)

reminder.SendReminders(google_form_name, google_form_url)

        