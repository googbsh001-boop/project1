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

    try:
        ws1 = sh.get_worksheet_by_id(0)
        
        print("--- Row 2 ~ 5 raw data ---")
        # I=9, J=10, K=11 in 1-indexed gspread. But let's just get the rows
        all_data = ws1.get_all_values()
        
        for r_idx in range(1, min(6, len(all_data))):
            row = all_data[r_idx]
            print(f"Row {r_idx+1} length: {len(row)}")
            if len(row) > 10:
                print(f"  I: '{row[8]}'")
                print(f"  J: '{row[9]}'")
                print(f"  K: '{row[10]}'")
                
                # Fetch raw numeric value using batch_get for these cells
                # to see if they are stored as numbers or strings 
                
    except Exception as e:
        print(f"디버그 에러: {e}")

if __name__ == "__main__":
    debug_sheet1()
