from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account

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
print(result["values"])
for row in result["values"]:
    print(row)


