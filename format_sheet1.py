import sys
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1n3WxFMxjS-mhHGE8I4dXi4Q2oJ3l4sq_OkCkBeJkbJI'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def format_sheet1():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)

    try:
        # gid=0 is usually the first sheet or the one created by default
        ws1 = sh.get_worksheet_by_id(0)
        
        # User requested: round to 3 decimal places, append '%' 
        # Example: 91.1234 -> 91.123%
        # The underlying value might be 0.911234 or 91.1234
        # We'll need to check the exact columns, let's assume it's column D,H,L,P like the other sheet, but we need to dynamically find it.
        # But this is '시트1' (ID:0) which might have '기초대비' in row 2 as well if it's identical to 입찰결과분석_상세
        
        headers = ws1.row_values(2)
        cols_to_format = []
        for i, h in enumerate(headers):
            if '기초대비' in h or '백분율' in h or '투찰율' in h or '낙찰율' in h:
                # 0-indexed i -> 1-indexed col
                from gspread.utils import rowcol_to_a1
                col_letter = rowcol_to_a1(1, i+1)[:-1]  # drop the '1'
                cols_to_format.append(col_letter)

        print(f"포맷할 열들: {cols_to_format}")
        
        # Setting the custom format (assuming values are e.g. 91.1234. If they are 0.911234, PERCENT type scales it * 100)
        # Assuming they are raw 91.1234 values since we used `0.0000"%"` earlier,
        # we'll use `0.000"%"` for 3 decimal places.
        fmt = CellFormat(numberFormat=NumberFormat(type='NUMBER', pattern='0.000"%"'))

        for c in cols_to_format:
            format_cell_range(ws1, f'{c}3:{c}1000', fmt)
        print("시트1 (ID: 0): 퍼센트 서식(소수점 셋째자리) 적용 완료!")
        
    except Exception as e:
        print(f"시트1 포맷팅 에러: {e}")

if __name__ == "__main__":
    format_sheet1()
