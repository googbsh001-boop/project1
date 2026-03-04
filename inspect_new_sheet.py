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

def main():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)
    
    print("Worksheets available in the new sheet:")
    for ws in sh.worksheets():
        print(f" - {ws.title} (ID: {ws.id})")
        if ws.id == 1426009222:
            try:
                headers = ws.row_values(1)
                print(f"   Headers: {headers}")
                if ws.row_count > 1:
                    row2 = ws.row_values(2)
                    print(f"   Row 2: {row2}")
            except Exception as e:
                print(f"   Could not fetch headers: {e}")

if __name__ == "__main__":
    main()
