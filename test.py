from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build

key = 'zmb54685508-c88132768091.json'

# Build 'Spreadsheet' object

spreadsheets_scope = [ 'https://www.googleapis.com/auth/spreadsheets' ]
sheets_credentials = ServiceAccountCredentials.from_json_keyfile_name(key, spreadsheets_scope)

sheets_service = build('sheets', 'v4', credentials=sheets_credentials)

# returns 'Spreadsheet' dict
# https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#resource-spreadsheet
spreadsheet = sheets_service.spreadsheets().create(
    body={
        "properties": {
            'title': 'spreadsheets test',
        },
        "sheets": [],
    }
).execute()


# id for the created file
spreadsheetId = spreadsheet['spreadsheetId']
# url of your file
spreadsheetUrl = spreadsheet['spreadsheetUrl']

# Build 'Permissions' object
drive_scope = [ 'https://www.googleapis.com/auth/drive' ]
drive_credentials = ServiceAccountCredentials.from_json_keyfile_name(key, drive_scope)

drive_service = build('drive', 'v3', credentials=drive_credentials)

# returns 'Permissions' dict
permissions = drive_service.permissions().create(
    fileId=spreadsheetId,
    transferOwnership=True,
    body={
        'type': 'user',
        'role': 'owner',
        'emailAddress': 'zhangmenbrewery@email.com',
    }
).execute()