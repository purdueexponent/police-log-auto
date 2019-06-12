# ryan chen (fatcat2)
# e: ryanjchen2@gmail.com

import xlrd
import google.auth
from googleapiclient.discovery import build
from dotenv import load_dotenv
import os

def main():
    # Let's initialize some values.
    # All env vars are located in the .env file (not in repo).
    load_dotenv()
    sheet_id = os.getenv('sheet_id')
    range_name = 'le_crime_log!A2:E'
    
    # Open the workbook in the source folder
    # TODO: make this work using input()
    wb = xlrd.open_workbook("src/wlpd_051419.xls")
    
    # Get the first and only sheet
    sheet = wb.sheets()[0]

    print(sheet.cell_value(0, 0))

    for row in sheet.get_rows():
        print(row)
    
    credentials, project = google.auth.default(
            scopes = ['https://www.googleapis.com/auth/spreadsheets']
    )

    print(project)
    
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()

    print(result)

if __name__ == '__main__':
    main()
