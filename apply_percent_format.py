import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1ZwmfkDFJDBLK_U2oymeI3XZ3w5FLLak7AxodXN1pXrU'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def format_percentages():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)

    # 1. 시트2 
    try:
        ws2 = sh.worksheet('시트2')
        # C열(투찰율) 서식 적용
        fmt = CellFormat(numberFormat=NumberFormat(type='PERCENT', pattern='0.0000%'))
        format_cell_range(ws2, 'C2:C1000', fmt)
        print("시트2: C열 퍼센트 포맷 적용 완료")
    except Exception as e:
        print(f"시트2 포맷팅 에러: {e}")

    # 2. 입찰결과분석_상세
    try:
        ws1 = sh.worksheet('입찰결과분석_상세')
        # In 입찰결과분석_상세, we check row 2 to find '기초대비(%)'.
        fmt2 = CellFormat(numberFormat=NumberFormat(type='NUMBER', pattern='0.0000"%"'))
        headers = ws1.row_values(2)
        cols_to_format = []
        for i, h in enumerate(headers):
            if '기초대비' in h:
                # 0-indexed i -> 1-indexed col
                from gspread.utils import rowcol_to_a1
                col_letter = rowcol_to_a1(1, i+1)[:-1]  # drop the '1'
                cols_to_format.append(col_letter)

        print(f"포맷할 열들: {cols_to_format}")
        for c in cols_to_format:
            format_cell_range(ws1, f'{c}3:{c}1000', fmt2)
        print("입찰결과분석_상세: 특정 열 퍼센트 기호 포맷 적용 완료")
    except Exception as e:
        print(f"입찰결과분석_상세 포맷팅 에러: {e}")

if __name__ == "__main__":
    format_percentages()
