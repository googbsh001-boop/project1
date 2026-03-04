import sys
import gspread
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1n3WxFMxjS-mhHGE8I4dXi4Q2oJ3l4sq_OkCkBeJkbJI'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def check_rows():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)
    ws1 = sh.get_worksheet_by_id(0)
    
    all_data = ws1.get_all_values()
    
    print("Row 77-79 Inspection:")
    for i in range(76, min(79, len(all_data))):
        row = all_data[i]
        b_val = row[1] if len(row) > 1 else ""
        c_val = row[2] if len(row) > 2 else ""
        print(f"Row {i+1}: B(공고명) = '{b_val}', C(입찰업체 갯수) = '{c_val}'")

if __name__ == '__main__':
    check_rows()
