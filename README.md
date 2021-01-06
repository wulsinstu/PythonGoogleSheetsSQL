# PythonGoogleSheetsSQL

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
