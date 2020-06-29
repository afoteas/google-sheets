from __future__ import print_function
from datetime import date
import pickle
import os.path
import numpy as np
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1cLvdj81ASDeCU1fq5jM4sBby6T1wxWv1FAuO7LDJlIQ'
SAMPLE_RANGE_NAME = 'test!A:F'
def init_service():
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
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    return service

def print_sheet():

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        i = 0
        for row in values:
            i = i+1
            if i == 1:
                # Print columns A and E, which correspond to indices 0 and 4.
                print('\n%s \t %s \t\t %s \t %s \t %s \t %s \t' % (row[0], row[1], row[2], row[3], row[4], row[5]))
                print("----------------------------------------------------------------------------")
            else:
                print('%s \t %s \t %s \t %s \t %s \t %s \t' % (row[0], row[1], row[2], row[3], row[4], row[5]))

def insert_sheet():
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    np_values = np.array(values)
    new_pos = int(max(np_values[1:,0])) + 1
    today = date.today()

    # dd/mm/YY
    today = today.strftime("%d/%m/%Y")

    first_name = input("First Name: ")
    last_name = input("Last Name: ")
    age = input("Age: ")
    amount = input("Amount: ")
    position = len(values) + 1
    values = [
        [
           new_pos, today, first_name, last_name, int(age), int(amount)
        ]
    ]
    body = {
        'values': values
    }

    range_input = 'A{}:F{}'.format(position,position)
    result = service.spreadsheets().values().update(
        spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_input, body=body, valueInputOption="RAW").execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))          

if __name__ == '__main__':
    service = init_service()
    while True:
        print("\n\nChoices: \n i) insert\n p) print\n e) exit\n t) test \n")
        choice = input("please select: ")
        if choice == 'i':
            insert_sheet()
        elif choice == 'p':
            print_sheet()
        elif choice == 'e':
            break
        elif choice == 't':
            insert_sheet()
        else:
            print("wrong choice, try again..")