import pandas as pd
from datetime import datetime
import warnings
import ssl # safety for emails
import smtplib  # for sending emails
import logging
from email.mime.multipart import MIMEMultipart  # for the messages in the emails
from email.mime.text import MIMEText
import openpyxl  # for excel files
import numpy as np # used to check for nan
import math # for is nan method
from email.message import EmailMessage
import sys

# ignores the annoying warnings
warnings.filterwarnings('ignore')

# setting up the logging file
logging.basicConfig(filename='failed_email_sends.log', level = logging.ERROR)

# opening both the Excel files
# has the user's names and applications
wb_lic = pd.read_excel('names_apps.xlsx', sheet_name=1, usecols='C:F')
# if names are missing from above workbook, we will use this one to extract the names
wb_names = pd.read_excel('computers.xlsx', sheet_name=1, usecols='B:D,H')

customerNames = []  # holds all the customer names
chargingRef = []    # holds all the customer charging references info
app = []            # holds all the customer application names
quantity = []       # holds all the corresponding quantities of apps
emails = []         # holds all the emails

# reading in from Excel file and populating lists
for index, row in wb_lic.iterrows():
   # if customer name is not available find it here and populate
   if pd.isna(wb_lic.loc[index, 'CustomerContact']):
      for namesIndex, namesRow in wb_names.iterrows():
         if namesRow['Asset'] == row['ChargingReferenceInformation']:
            fullCustomerContactName = namesRow['Last Name'] + " " + namesRow['First Name'] + " (" + namesRow['Dept'] + ")"
            customerNames.append(fullCustomerContactName)

            # get email and append it
            email = namesRow['First Name'] + "." + namesRow['Last Name'] + "@bosch.com"
            emails.append(email)

            # if the customer name is not empty then append the rest
            if not fullCustomerContactName == None:
               chargingRef.append(row['ChargingReferenceInformation'])
               app.append(row['ApplicationName'])
               quantity.append(row['Quantity'])

   # if name is available, append it straight from the dataframe
   else:
      customerNames.append(row['CustomerContact'])
      chargingRef.append(row['ChargingReferenceInformation'])
      app.append(row['ApplicationName'])
      quantity.append(row['Quantity'])


# the full dataframe with all the missing customer names
df = pd.DataFrame(
   {'ChargingReferenceInformation': chargingRef,
    'CustomerContact': customerNames,
    'ApplicationName': app,
    'Quantity': quantity
   })



# CREATING MESSAGES FOR EVERY CUSTOMER

# default message that will go on top of every email
defaultmessage = '<tr style ="font-size:18px"><td>{:<40}</td> <td>{:<50}</td> <td>{:<40}</td> <td>{:<10}</td></tr>'.format(
   'Charging Reference Info', 'Customer Contact', 'Application', 'Quantity'
)

# dictionary that will hold the customer and their own message
customer_and_message = {}

# list that holds the customer email and app
customers = []

for index, row in df.iterrows():
   # getting customer app
   customer_app = df.loc[index, 'ApplicationName']
   # getting the customer email
   customer_name = df.loc[index, 'CustomerContact']
   customer_list = customer_name.split()
   customer_email = customer_list[1] + '.' + customer_list[0] + '@us.bosch.com'

   # if new customer, add the customer and their message to the dictionary
   if customer_email not in [person[0] for person in customers]:
      # adding the default message and their info
      customer_and_message[customer_email] = defaultmessage
      customer_and_message[customer_email] += '<tr><td>{:<40}</td> <td>{:<50}</td> <td>{:<40}</td> <td>{:<10}</td></tr>'.format(
            df.loc[index, 'ChargingReferenceInformation'], df.loc[index, 'CustomerContact'],
            df.loc[index, 'ApplicationName'], df.loc[index, 'Quantity']
         )

   # if customer previously added and the app is different, update their message with new app info
   if customer_email in [person[0] for person in customers] and customer_app not in [app[1] for app in customers if app[0] == customer_email]:
      customer_and_message[customer_email] += '<tr><td>{:<40}</td> <td>{:<50}</td> <td>{:<40}</td> <td>{:<10}</td></tr>'.format(
         df.loc[index, 'ChargingReferenceInformation'], df.loc[index, 'CustomerContact'],
         df.loc[index, 'ApplicationName'], df.loc[index, 'Quantity']
      )

   # adding the customer and their app to the list, so it can be checked on the next loops
   customers.append((customer_email, df.loc[index, 'ApplicationName']))



# EMAILING
sender_email = 'whatever@bosch.com' # changed for security
subject = 'Response Required: Do you utilize the following applications'

for key, value in customer_and_message.items():
   msg = MIMEMultipart('alternative')
   msg['From'] = sender_email
   msg['To'] = key
   msg['Subject'] = subject

   email_string = '''<h2>Are you still utilizing the following applications?</h2>
   <style>
      td {
        padding-right:50px;
      }
   </style>
   <table class = "data">'''
   email_string += value
   email_string += '</table> <br><h4>Applications will be removed if no response for cost saving purposes.</h4>'

   html = MIMEText(email_string, 'html')
   msg.attach(html)
   smtp = smtplib.SMTP('bosch account', 555) # changed for security purposes
   smtp.sendmail(msg['From'], msg['To'], msg.as_string())

print("Successfully sent emails")
