from __future__ import print_function

import os.path
import time

from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SPREADSHEET_ID = '1QNLONctqPKp48HpnbrkAEe49HIfQfucZTQsx9WFacUs'
RANGE_NAME = 'Розклад!C300:D301'
DATA_TO_WRITE = [['Марк', '662272536'], ['Марк', '662272536']]

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        try:
            while True:
                run_time = datetime.strptime('11:16:58.2', '%H:%M:%S.%f')
                now_time = datetime.now().strftime("%H:%M:%S.%f")
                if now_time >= run_time.strftime("%H:%M:%S.%f"):
                    print('start', now_time)
                    write_to_sheet(service, SPREADSHEET_ID, RANGE_NAME, DATA_TO_WRITE)
                    break
                # else:
                    # print('rano')
                time.sleep(.05)
                # print(now_time)
                
        except HttpError as error:
            print(f"Произошла ошибка: {error}")
            print("Скрипт завершен.")
    except HttpError as err:
        print(err)

def write_to_sheet(service, spreadsheet_id, range_name, values):
    body = {
        'values': values
    }

    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=range_name,
        valueInputOption='RAW', body=body).execute()

    now_time = datetime.now().strftime("%H:%M:%S.%f")
    print(f'Обновлено {result.get("updatedCells")} ячеек.')
    print('end', now_time)


if __name__ == '__main__':
    main()