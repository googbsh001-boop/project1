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

def format_sheet2_row_g():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)

    try:
        ws = sh.get_worksheet_by_id(WORKSHEET_ID)
        
        print("시트2 (ID: 1426009222): 모든 데이터를 가져옵니다...")
        all_data = ws.get_all_values()
        
        if len(all_data) <= 1:
            print("데이터가 없습니다.")
            return

        num_cols = len(all_data[0])
        max_col_letter = gspread.utils.rowcol_to_a1(1, num_cols)[:-1]

        # 사용자가 "G열의 값이 5 이하인 행만 채우기해줘" 라고 지정함.
        # 기존 E열(한화건설 순위) 하이라이트 조건과 차별성을 두기 위해 다른 연한 색상(예: 연한 하늘색)을 지정
        highlight_color = Color(0.85, 0.90, 0.98) 
        fmt_highlight = CellFormat(backgroundColor=highlight_color, textFormat=TextFormat(bold=False))
        
        batch = []
        count = 0
        
        # G열은 index 6 (0부터 시작하므로)
        for r_idx in range(1, len(all_data)):
            row = all_data[r_idx]
            if len(row) > 6:
                val_str = row[6] # G column (HDC 낙찰우선순위)
                if val_str:
                    try:
                        val_float = float(val_str.strip())
                        if val_float <= 5.0:
                            row_num = r_idx + 1
                            cell_range = f'A{row_num}:{max_col_letter}{row_num}'
                            batch.append((cell_range, fmt_highlight))
                            count += 1
                    except ValueError:
                        pass
                        
        if batch:
            print(f"G열 값이 5 이하인 행 총 {count}건에 하이라이트 서식(연한 하늘색)을 적용합니다...")
            format_cell_ranges(ws, batch)
            print("행 색상 채우기 완료!")
        else:
            print("5 이하인 행을 찾지 못했습니다.")

    except Exception as e:
        print(f"시트2 포맷팅 에러: {e}")

if __name__ == "__main__":
    format_sheet2_row_g()
