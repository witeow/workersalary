from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
import numpy as num
import pandas as pd

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'google_excel_key.json'

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1b4UOb4PrexdcJTyLv1ptjL3n8tfTYV4HAlX20x-Zc6Y'


service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                            range="workinghours").execute()
values = result.get('values', [])
# print(result)

# print(len(values)) 

# creating a dictionary with location as key and name as value
# for example:
# 55_lentor_way = {"Name" : [Islam, Subrot]
# "1" : [8,4,8,0]
# "2" : [8,2,8,3]}
###################################################################################
# creating different locations for different sheets
locations = []
for row in values[1:]:
        try:
                if row[3] not in locations:
                        locations.append(row[3])
                if row[5] not in locations:
                        locations.append(row[5])
        except IndexError:
                continue

print(locations)


