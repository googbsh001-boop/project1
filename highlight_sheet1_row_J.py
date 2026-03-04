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

def format_sheet1_row_j():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)

    try:
        ws1 = sh.get_worksheet_by_id(0)
        
        print("시트1 (ID: 0): 모든 데이터를 가져옵니다...")
        all_data = ws1.get_all_values()
        
        if len(all_data) <= 1:
            print("데이터가 없습니다.")
            return

        # Header 갯수 파악하여 끝 열 알아내기 (보통 M열=13까지 있음)
        num_cols = len(all_data[0])
        max_col_letter = gspread.utils.rowcol_to_a1(1, num_cols)[:-1]

        # J열 = 9 (0-indexed)
        # 88.000% 이상인 경우 하이라이트.
        # 밝은 파란색이나 연한 주황색 등 강조 색상을 지정해줍니다. 여기선 연한 주황색(또는 노란색)으로 해보겠습니다.
        highlight_color = Color(1.0, 0.95, 0.8) # Light yellow/orange
        fmt_highlight = CellFormat(backgroundColor=highlight_color, textFormat=TextFormat(bold=True))
        
        batch = []
        count = 0
        
        # 2번째 행부터 순회
        for r_idx in range(1, len(all_data)):
            row = all_data[r_idx]
            
            if len(row) > 9:
                val_str = row[9] # J column
                if val_str:
                    try:
                        clean_val = val_str.replace('%', '').replace(',', '').strip()
                        if clean_val:
                            val_float = float(clean_val)
                            # 값은 이미 88.000 형태이므로 88.0 이상인지 비교
                            if val_float >= 88.0:
                                # A열부터 해당 데이터 길이만큼의 열까지 색상 적용
                                # (e.g., A2:M2)
                                row_num = r_idx + 1
                                cell_range = f'A{row_num}:{max_col_letter}{row_num}'
                                batch.append((cell_range, fmt_highlight))
                                count += 1
                    except ValueError:
                        pass
                        
        if batch:
            print(f"J열 값이 88.000% 이상인 행 총 {count}건에 하이라이트 서식을 적용합니다...")
            format_cell_ranges(ws1, batch)
            print("행 색상 채우기 완료!")
        else:
            print("88.000% 이상인 행을 찾지 못했습니다.")

    except Exception as e:
        print(f"시트1 포맷팅 에러: {e}")

if __name__ == "__main__":
    format_sheet1_row_j()
