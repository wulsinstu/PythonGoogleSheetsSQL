# -*- coding: utf-8 -*-
"""
This file was develped using instructions from these two websites:
    https://developers.google.com/sheets/api/quickstart/python
    https://medium.com/analytics-vidhya/how-to-read-and-write-data-to-google-spreadsheet-using-python-ebf54d51a72c

A. This was written to be run in the Anaconda Python environment. 
It should also work in other environments, but you'll have to figure out how to install the proper libraries.

B. Before using this you'll need to install some libaries. Open Anaconda Prompt and type these commands:
pip install SQLAlchemy
pip install pandas
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

C. You'll need to turn on the Google Sheets API using the button in step one on this page: https://developers.google.com/sheets/api/quickstart/python
  In resulting dialog click DOWNLOAD CLIENT CONFIGURATION and save the file credentials.json to your working directory. 
  
D. Update the variables with your Google Sheet and SQL info 

"""
## Libraries--------------------------------------------------------------------------
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import sqlalchemy
import pandas

## Google Variables--------------------------------------------------------------------------
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SPREADSHEET_ID = '1aS96_zQOBwawnBxdxoA50KXt-pHItN5HJz_upm2VInw' #This is pulled from the middle part of the url
RANGE_NAME = 'Sheet1!A1:G100'


## SQL Variables ----------------------------------------------------------------------------
user = 'swulsin'
passw = os.getenv('IC_SQL_PASSWORD') #You can also just type in your password here if you don't know how to use Environment Variables
host = 'elhaynes.infinitecampus.org'
port = '7771'
DATABASE_NAME = 'elhaynes'
SCHEMA = 'elhcustom'
table = 'tmpGoogleSync'

## Main Program ------------------------------------------------------------------------------
## 1. Get data from the Google Sheet----------------------------------------------------------
print('Getting data from ' + RANGE_NAME)
creds = None 
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)
service = build('sheets', 'v4', credentials=creds)
# Call the Sheets API
sheet = service.spreadsheets()
# result pulls the data from the spreadsheet
result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
# values saves the data as a dictionary
values = result.get('values', [])

## 2. Use Pandas to convert values to a data frame ---------------------------------------------------
# frame converts values to a Pandas dataframe with columns and rows. Row 1 is used for the column names
frame = pandas.DataFrame(values, columns = values[0])
# frame2 removes the first row, which repeats the column headers
frame2 = frame.drop(0)
print(frame2)

## 3. Save Data Frame to SQL ----------------------------------------------------------------------------
# Connect to the SQL Database
engine = sqlalchemy.create_engine('mssql+pyodbc://{}:{}@{}:{}/{}?driver=ODBC+Driver+17+for+SQL+Server'.format(user, passw, host, port, DATABASE_NAME))
connection = engine.connect()
metadata = sqlalchemy.MetaData()
print('connection started')

#Copy data from frame2 to SQL table
frame2.to_sql(table, con=engine, schema = SCHEMA, if_exists='replace')
print('data sent to SQL table ' + table)
