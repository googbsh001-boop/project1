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

def clear_empty_b_matches():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)
    ws1 = sh.get_worksheet_by_id(0)
    
    all_data = ws1.get_all_values()
    
    if len(all_data) > 0 and '입찰업체 갯수' in all_data[0]:
        col_to_update = all_data[0].index('입찰업체 갯수')
    else:
        return

    updates = []
    
    for r_idx in range(1, len(all_data)):
        row = all_data[r_idx]
        b_val = row[1].strip() if len(row) > 1 else ""
        
        if b_val == "":
            updates.append({
                'range': gspread.utils.rowcol_to_a1(r_idx + 1, col_to_update + 1),
                'values': [[""]]
            })
            
    if updates:
        print(f"B열(공고명)이 비어있는 {len(updates)}개 행의 잘못된 카운트(19)를 지웁니다.")
        ws1.batch_update(updates, value_input_option='USER_ENTERED')
        print("지우기 완료!")
    else:
        print("정리할 내용이 없습니다.")

if __name__ == '__main__':
    clear_empty_b_matches()
