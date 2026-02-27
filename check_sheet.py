import gspread
from google.oauth2.service_account import Credentials
import os

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1ICIoAj-KU-BeYKFEvGNXtpCxpUOXa-JS4Kg0Zy_UOFw'
WORKSHEET_NAME = '리스트'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    credentials = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.worksheet(WORKSHEET_NAME)
    
    # Read the data to see what columns we have
    data = worksheet.get_all_values()
    print("Total rows:", len(data))
    if len(data) > 0:
        print("Header:", data[0])
    for i in range(1, min(6, len(data))):
        print(f"Row {i+1}:", data[i])
        
if __name__ == "__main__":
    main()
