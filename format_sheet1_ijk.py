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

def format_sheet1_ijk():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)

    try:
        ws1 = sh.get_worksheet_by_id(0)
        
        print("시트1 (ID: 0): I, J, K열의 데이터를 가져옵니다...")
        
        all_data = ws1.get_all_values()
        if len(all_data) <= 1:
            print("업데이트할 데이터가 없습니다.")
            return
            
        updates = []
        for r_idx in range(1, len(all_data)):
            row = all_data[r_idx]
            while len(row) <= 10:
                row.append("")
                
            for c_idx in [8, 9, 10]:
                val_str = row[c_idx]
                if val_str:
                    try:
                        # 현 구글 데이터는 문자열 9889.830% 등으로 나옵니다.
                        # clean_val = 9889.83
                        clean_val = val_str.replace(',', '').replace('%', '').strip()
                        if clean_val:
                            val_float = float(clean_val)
                            
                            # val_float는 현재 눈에 보이는 9889.830 등의 숫자입니다.
                            # 여기서 사용자가 원하는 표시는 98.898% 입니다.
                            # PERCENT 형식에 98.898%로 보이려면 내부 숫자는 0.98898 이어야 합니다.
                            # 따라서 현재 값(9889.830)에서 10000으로 나누어야 0.988983이 됩니다.
                            new_val = val_float / 10000.0
                            
                            updates.append({
                                'range': gspread.utils.rowcol_to_a1(r_idx + 1, c_idx + 1),
                                'values': [[new_val]]
                            })
                    except ValueError:
                        pass
                        
        if updates:
            print(f"총 {len(updates)}개의 셀 값을 현재 보이는 값 기준으로 100을 나누도록 조정하여 업데이트합니다...")
            ws1.batch_update(updates, value_input_option='USER_ENTERED')
            print("데이터 업데이트 완료!")
        else:
            print("업데이트할 숫자 데이터가 없습니다.")

        # Re-apply Percentage format
        fmt = CellFormat(numberFormat=NumberFormat(type='PERCENT', pattern='0.000%'))
        format_cell_range(ws1, 'I2:K1000', fmt)
        print("시트1 (ID: 0): I, J, K열 퍼센트 서식 재적용 완료!")
        
    except Exception as e:
        print(f"시트1 포맷팅 에러: {e}")

if __name__ == "__main__":
    format_sheet1_ijk()
