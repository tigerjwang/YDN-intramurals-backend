from __future__ import print_function

import csv
import os.path
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
STANDINGS_SPREADSHEET_ID = '1xffq-cDUbEzLkenPLQDPCz9U9qfUWMqPBDUUHg-Ak0Y'

DEBUG = True

COLLEGE_NAME = {
    "SY": "Saybrook",
    "MY": "Pauli Murray",
    "TD": "Timothy Dwight",
    "MC": "Morse",
    "GH": "Grace Hopper",
    "BF": "Benjamin Franklin",
    "DC": "Davenport",
    "ES": "Ezra Stiles",
    "TC": "Trumbull",
    "PC": "Pierson",
    "SM": "Silliman",
    "BK": "Berkeley",
    "BR": "Branford",
    "JE": "Jonathan Edwards",
}

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
    row_start = 2
    row_end = 15
    college_range = f'{sheet_name}!{college_column[sheet_name]}{row_start}:{college_column[sheet_name]}{row_end}'
    score_range = f'{sheet_name}!{score_column[sheet_name]}{row_start}:{score_column[sheet_name]}{row_end}'

    # Invoke Google Sheets API
    try:
        college_result = sheet.values().get(spreadsheetId=STANDINGS_SPREADSHEET_ID,
                                            range=college_range).execute().get('values', [])
        score_result = sheet.values().get(spreadsheetId=STANDINGS_SPREADSHEET_ID,
                                            range=score_range).execute().get('values', [])
    except HttpError as err:
        print(err)
    
    # Output results to CSV
    if college_result and score_result:
        if DEBUG: print("college_result:", college_result)
        if DEBUG: print("score_result:", score_result)
        with open(f'standings_{sheet_name.lower()}.csv', 'w', newline='') as csvfile:
            fieldnames = ['college', 'score']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for i in range(row_end - row_start + 1):
                college_abbreviated = college_result[i][0]
                college_name = COLLEGE_NAME[college_abbreviated]
                writer.writerow({'college': college_name, 'score': score_result[i][0]})
    else:
        print('No data found on college or score column.')


def main():
    if (len(sys.argv) != 2):
        print("Usage: ./scraper API_KEY")
        return

    api_key = sys.argv[1]
    service = build('sheets', 'v4', developerKey=api_key)

    sheet = service.spreadsheets()

    scrape_standings(sheet, "overall")
    scrape_standings(sheet, "WHoops")
    scrape_standings(sheet, "MHoops")
    scrape_standings(sheet, "BHoops")
    scrape_standings(sheet, "CHoops")
    scrape_standings(sheet, "Volleyball")
    scrape_standings(sheet, "Dodgeball")


if __name__ == '__main__':
    main()
