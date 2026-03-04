import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1n3WxFMxjS-mhHGE8I4dXi4Q2oJ3l4sq_OkCkBeJkbJI'
BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def extract_bidder_count_dynamic(base_dir):
    bidder_counts = {}
    
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.endswith('.xlsb') and not file.startswith('~'):
                file_path = os.path.join(root, file)
                print(f"엑셀 읽는 중: {file}")
                
                try:
                    try:
                        df = pd.read_excel(file_path, engine='pyxlsb', sheet_name='입찰결과', header=None)
                    except Exception as e:
                        df = pd.read_excel(file_path, engine='pyxlsb', header=None)
                    
                    header_row_idx = -1
                    col_idx_company = -1
                    
                    # '회사명' 셀을 찾습니다. 사용자가 'B열의 "회사명"이라고 입력된 칸의 아랫칸부터' 라고 했습니다.
                    # B열은 주로 index 1입니다. 전체 데이터프레임을 순회하며 '회사명' 텍스트를 찾습니다.
                    for r_idx, row in df.iterrows():
                        for c_idx, cell_val in enumerate(row):
                            if str(cell_val).replace(" ", "") == "회사명":
                                header_row_idx = r_idx
                                col_idx_company = c_idx
                                break
                        if header_row_idx != -1:
                            break
                            
                    if header_row_idx != -1 and col_idx_company != -1:
                        # 회사명 아래 행부터 끝까지 해당 열의 데이터 추출
                        col_data = df.iloc[header_row_idx+1:, col_idx_company]
                        
                        valid_bidders = col_data.dropna().astype(str).str.strip()
                        # 빈 문자열, nan, "-", "0" 등의 값을 필터링
                        valid_bidders = valid_bidders[
                            (valid_bidders != '') & 
                            (valid_bidders.str.lower() != 'nan') & 
                            (valid_bidders != '-') &
                            (valid_bidders != '0') &
                            (valid_bidders != '0.0') &
                            (valid_bidders.str.len() > 0)
                        ]
                        count = len(valid_bidders)
                        print(f"  -> {file} 참여업체: {count}개 추출 (헤더 인덱스: {header_row_idx}, 컬럼 인덱스: {col_idx_company})")
                    else:
                        print("  -> '회사명' 셀을 찾을 수 없습니다.")
                        count = 0
                    
                    key = file.replace('.xlsb', '').strip()
                    bidder_counts[key] = count
                    
                except Exception as e:
                    print(f"  -> 파일 처리 중 에러: {e}")
                    
    return bidder_counts

def update_col_in_sheet1():
    print("업체 수 데이터(회사명 아랫칸부터, '-' 등 제외) 카운트 시작...")
    bidder_counts = extract_bidder_count_dynamic(BASE_DIR)
    
    client = gspread.authorize(Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES))
    sh = client.open_by_key(SHEET_ID)
    ws1 = sh.get_worksheet_by_id(0)
    
    all_data = ws1.get_all_values()
    
    if len(all_data) > 0 and '입찰업체 갯수' in all_data[0]:
        col_to_update = all_data[0].index('입찰업체 갯수')
        print(f"'{all_data[0][col_to_update]}' 열에 새로운 카운트로 업데이트를 진행합니다. (인덱스: {col_to_update})")
    else:
        print("'입찰업체 갯수' 열을 찾을 수 없습니다.")
        return

    updates = []
    for r_idx in range(1, len(all_data)):
        row = all_data[r_idx]
        if len(row) > 1: # B열 공고명이 있는 경우
            b_val = row[1].strip()
            matched_count = ""
            for key, count in bidder_counts.items():
                if b_val in key or key in b_val:
                    matched_count = count
                    break
            
            if matched_count != "":
                # 업데이트 값 큐
                updates.append({
                    'range': gspread.utils.rowcol_to_a1(r_idx + 1, col_to_update + 1),
                    'values': [[matched_count]]
                })
                
    if updates:
        print(f"총 {len(updates)}건의 매칭된 업체 수를 갱신합니다...")
        ws1.batch_update(updates, value_input_option='USER_ENTERED')
        print("업데이트 완료!")
    else:
        print("수정할 데이터가 없습니다.")

if __name__ == '__main__':
    update_col_in_sheet1()
