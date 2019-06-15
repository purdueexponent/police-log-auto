# ryan chen (fatcat2)
# e: ryanjchen2@gmail.com
import xlrd
import google.auth
from googleapiclient.discovery import build
from dotenv import load_dotenv
import os


# Method to abstract the appending of the values to the gSheet.
def write_to_sheet(sheet, sheet_id, range_name, value_input_option, body):
    # Google Python API Client method to write values
    sheet.values().append(
            spreadsheetId=sheet_id,
            range=range_name,
            valueInputOption="USER_ENTERED",
            body=body
    ).execute()


# Method to convert the xlrd rows into arrays while processing
# xlrdate objects.
def processRow(row):
    tmplist = []
    for cell in row:
        if(cell.ctype == 3):
            tmplist.append(
                    str(
                        xlrd.xldate.xldate_as_datetime(
                            cell.value, 0
                            ).date()
                        )
                    )
        else:
            tmplist.append(cell.value)
    return tmplist


def main(*argv):
    # Let's initialize some values.
    # All env vars are located in the .env file (not in repo).
    load_dotenv()
    sheet_id = os.getenv('sheet_id')
    range_name = 'le_crime_log!A2:S'

    # Open the workbook in the source folder
    # TODO: make this work using input()
    whereString = input("Please input folder path: ")

    wb = xlrd.open_workbook(whereString, formatting_info=True)

    # Get the first and only sheet of the Excel sheet
    sheet = wb.sheets()[0]

    # Get credentials to write to Google Sheets API
    credentials, project = google.auth.default(
            scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    service = build('sheets', 'v4', credentials=credentials)

    # Create gsheet object
    gsheet = service.spreadsheets()

    # Import values of Excel sheet
    values = []

    for row in sheet.get_rows():
        values.append(processRow(row))

    # Pop the first list since those are the headers
    values.pop(0)

    # Write the new values to the sheet
    body = {
        'values': values
    }

    write_to_sheet(gsheet, sheet_id, range_name, "USER_INPUT", body)


if __name__ == '__main__':
    main()
