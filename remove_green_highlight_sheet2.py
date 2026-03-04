import sys
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1n3WxFMxjS-mhHGE8I4dXi4Q2oJ3l4sq_OkCkBeJkbJI'
WORKSHEET_ID = 1426009222
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def remove_green_and_reapply_blue():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)

    try:
        ws = sh.get_worksheet_by_id(WORKSHEET_ID)
        
        print("시트2 (ID: 1426009222): A2:G1000 배경색을 흰색으로 초기화하여 기존 녹색 하이라이트를 제거합니다...")
        # Number format 등 다른 속성은 건드리지 않고 배경색만 흰색으로 덮어씁니다.
        white_format = CellFormat(backgroundColor=Color(1.0, 1.0, 1.0))
        format_cell_range(ws, 'A2:G1000', white_format)

        print("HDC(G열) 5 이하인 행에 셀 배경색(연한 하늘색)을 다시 확인 및 복구합니다...")
        all_data = ws.get_all_values()
        if len(all_data) <= 1:
            print("데이터가 없습니다.")
            return

        num_cols = len(all_data[0])
        max_col_letter = gspread.utils.rowcol_to_a1(1, num_cols)[:-1]

        highlight_blue = Color(0.85, 0.90, 0.98) 
        fmt_blue = CellFormat(backgroundColor=highlight_blue, textFormat=TextFormat(bold=False))
        
        batch = []
        count = 0
        for r_idx in range(1, len(all_data)):
            row = all_data[r_idx]
            if len(row) > 6:
                val_str = row[6] # G column
                if val_str:
                    try:
                        val_float = float(val_str.strip())
                        if val_float <= 5.0:
                            row_num = r_idx + 1
                            cell_range = f'A{row_num}:{max_col_letter}{row_num}'
                            batch.append((cell_range, fmt_blue))
                            count += 1
                    except ValueError:
                        pass

        if batch:
            format_cell_ranges(ws, batch)
            print(f"행 색상 채우기(연한 하늘색) {count}건 업데이트 완료!")
        else:
            print("HDC 5 이하인 행을 찾지 못했습니다.")
            
    except Exception as e:
        print(f"시트2 포맷팅 에러: {e}")

if __name__ == "__main__":
    remove_green_and_reapply_blue()
