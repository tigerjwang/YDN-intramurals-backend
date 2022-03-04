from __future__ import print_function

import csv
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

API_KEY = 'AIzaSyDsmqiCdJ2z36NFseEizttPecJHraQaLAE'

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
STANDINGS_SPREADSHEET_ID = '1xffq-cDUbEzLkenPLQDPCz9U9qfUWMqPBDUUHg-Ak0Y'
SAMPLE_RANGE_NAME = 'overall!J2:J15'

# class Sheet(Enum):
#     OVERALL
#     WHOOPS
#     MHOOPS
#     BHOOPS
#     CHOOPS
#     VOLLEYBALL
#     DODGEBALL


def scrape_standings(sheet, sheet_name):
    college_column = {
        "overall": "A",
        "WHoops": "F",
        "MHoops": "F",
        "BHoops": "F",
        "CHoops": "F",
        "Volleyball": "F",
        "Dodgeball": "F"
    }
    score_column = {
        "overall": "J",
        "WHoops": "K",
        "MHoops": "J",
        "BHoops": "J",
        "CHoops": "J",
        "Volleyball": "J",
        "Dodgeball": "J"
    }
    college_range = f'{sheet_name}!{college_column[sheet_name]}2:{college_column[sheet_name]}15'
    score_range = f'{sheet_name}!{score_column[sheet_name]}2:{score_column[sheet_name]}15'
    try:
        college_result = sheet.values().get(spreadsheetId=STANDINGS_SPREADSHEET_ID,
                                            range=college_range).execute().get('values', [])
        score_result = sheet.values().get(spreadsheetId=STANDINGS_SPREADSHEET_ID,
                                            range=score_range).execute().get('values', [])
        if college_result and score_result:
            print(college_result)
            print(score_result)
            # with open('names.csv', 'w', newline='') as csvfile:
            #     fieldnames = ['first_name', 'last_name']
            #     writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            #     writer.writeheader()
            #     writer.writerow({'first_name': 'Baked', 'last_name': 'Beans'})
            #     writer.writerow({'first_name': 'Lovely', 'last_name': 'Spam'})
            #     writer.writerow({'first_name': 'Wonderful', 'last_name': 'Spam'})
        else:
            print('No data found on college or score column.')
    except HttpError as err:
        print(err)
    
    

def main():
    service = build('sheets', 'v4', developerKey=API_KEY)

    # Call the Sheets API
    sheet = service.spreadsheets()

    scrape_standings(sheet, "WHoops")


if __name__ == '__main__':
    main()