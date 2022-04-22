from __future__ import print_function
from collections import namedtuple

import csv
import os.path
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

Spreadsheet = namedtuple('Spreadsheet', 'id year season')
SPREADSHEETS = [
    Spreadsheet('1xffq-cDUbEzLkenPLQDPCz9U9qfUWMqPBDUUHg-Ak0Y', 2021, 'winter'),
    Spreadsheet('1wLYwmOjzGjFvZInKKpCGcqNerMzU8VzE5cxskipAudw', 2021, 'fall'),
    Spreadsheet('19V_NTYYFbBQXVFJumDHMaF9OQUUC9ZbQxyZJnUQBQ9U', 2022, 'spring'),
]
SHEET_NAMES = {
    'winter': ['overall', 'WHoops', 'MHoops', 'BHoops', 'CHoops', 'Volleyball', 'Dodgeball'],
    'fall': ['overall', 'spikeball', 'corn', 'tabletennis', 'pickle', 'kan', 'football', 'soccer'],
    'spring': ['overall', 'IndoorSoccer', 'KanJam', 'Kickball', 'Badminton'],
}

DEBUG = True

COLLEGE_NAME = {
    'SY': 'Saybrook',
    'MY': 'Pauli Murray',
    'TD': 'Timothy Dwight',
    'MC': 'Morse',
    'GH': 'Grace Hopper',
    'BF': 'Benjamin Franklin',
    'DC': 'Davenport',
    'ES': 'Ezra Stiles',
    'TC': 'Trumbull',
    'PC': 'Pierson',
    'SM': 'Silliman',
    'BK': 'Berkeley',
    'BR': 'Branford',
    'JE': 'Jonathan Edwards',
}

def scrape_standings(sheet, spreadsheet, sheet_name):
    if DEBUG: print('running scraper for', spreadsheet.year, spreadsheet.season, sheet_name)
    college_column = {
        'overall': 'A',
        'WHoops': 'F',
        'MHoops': 'F',
        'BHoops': 'F',
        'CHoops': 'F',
        'Volleyball': 'F',
        'Dodgeball': 'F',

        'spikeball': 'F',
        'corn': "F",
        'tabletennis': 'F',
        'pickle': 'F',
        'kan': 'F',
        'football': 'F',
        'soccer': 'F',

        'IndoorSoccer': 'F',
        'KanJam': 'F',
        'Kickball': 'F',
        'Badminton': 'F',
    }
    score_column = {
        'overall': 'J',
        'WHoops': 'K',
        'MHoops': 'J',
        'BHoops': 'J',
        'CHoops': 'J',
        'Volleyball': 'J',
        'Dodgeball': 'J',

        'spikeball': 'K',
        'corn': "J",
        'tabletennis': 'J',
        'pickle': 'J',
        'kan': 'J',
        'football': 'J',
        'soccer': 'J',

        'IndoorSoccer': 'K',
        'KanJam': 'J',
        'Kickball': 'J',
        'Badminton': 'J',
    }
    row_start = 2
    row_end = 15
    college_range = f'{sheet_name}!{college_column[sheet_name]}{row_start}:{college_column[sheet_name]}{row_end}'
    score_range = f'{sheet_name}!{score_column[sheet_name]}{row_start}:{score_column[sheet_name]}{row_end}'

    # Invoke Google Sheets API
    try:
        college_result = sheet.values().get(spreadsheetId=spreadsheet.id,
                                            range=college_range).execute().get('values', [])
        score_result = sheet.values().get(spreadsheetId=spreadsheet.id,
                                            range=score_range).execute().get('values', [])
    except HttpError as err:
        print(err)
    
    if college_result and score_result:
        if DEBUG: print('college_result:', college_result)
        if DEBUG: print('score_result:', score_result)

        result = []
        for i in range(row_end - row_start + 1):
            college_abbreviated = college_result[i][0]
            college_name = COLLEGE_NAME[college_abbreviated]
            result.append({'college': college_name, 'score': score_result[i][0]})
        
        # Sort results
        result.sort(reverse=True, key=lambda entry: float(entry['score']))

        # Output results to CSV
        with open(f'standings_{sheet_name.lower()}_{spreadsheet.year}_{spreadsheet.season}.csv', 'w', newline='') as csvfile:
            fieldnames = ['college', 'score']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for i in range(row_end - row_start + 1):
                writer.writerow({'college': result[i]['college'], 'score': result[i]['score']})
    else:
        print('No data found on college or score column.')


def main():
    if (len(sys.argv) != 2):
        print('Usage: ./scraper API_KEY')
        return

    api_key = sys.argv[1]
    service = build('sheets', 'v4', developerKey=api_key)

    sheet = service.spreadsheets()

    for spreadsheet in SPREADSHEETS:
        for sheet_name in SHEET_NAMES[spreadsheet.season]:
            scrape_standings(sheet, spreadsheet, sheet_name)


if __name__ == '__main__':
    main()
