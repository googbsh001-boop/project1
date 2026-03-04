import sys
import gspread
from google.oauth2.service_account import Credentials

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1n3WxFMxjS-mhHGE8I4dXi4Q2oJ3l4sq_OkCkBeJkbJI'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def debug_sheet1():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)
    ws1 = sh.get_worksheet_by_id(0)
    
    all_data = ws1.get_all_values()
    print("Row 0 (Headers):", all_data[0][:5])
    
    for i in range(1, min(10, len(all_data))):
        print(f"Row {i}: B='{all_data[i][1]}', C='{all_data[i][2] if len(all_data[i])>2 else ''}'")

if __name__ == '__main__':
    debug_sheet1()
