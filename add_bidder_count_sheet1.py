import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
from gspread_formatting import *

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1n3WxFMxjS-mhHGE8I4dXi4Q2oJ3l4sq_OkCkBeJkbJI'
BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_google_sheet_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def extract_bidder_count(base_dir):
    # 각 공고명(또는 파일명)에 매칭되는 입찰자 수를 저장하는 딕셔너리
    bidder_counts = {}
    
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.endswith('.xlsb') and not file.startswith('~'):
                file_path = os.path.join(root, file)
                print(f"엑셀 읽는 중: {file}")
                
                try:
                    df = pd.read_excel(file_path, engine='pyxlsb', header=None)
                    
                    header_row_idx = -1
                    for idx, row in df.iterrows():
                        if str(row[0]).replace(" ", "") == "순위" and str(row[1]).replace(" ", "") == "회사명":
                            header_row_idx = idx
                            break
                    
                    if header_row_idx == -1:
                        print(f"  -> 헤더 행을 찾을 수 없습니다.")
                        continue
                        
                    data_df = df.iloc[header_row_idx+1:]
                    
                    # 입찰에 참여한 업체 수 계산 (결측치가 아닌 데이터 길이)
                    # 순위 열(보통 0번 인덱스)에 데이터가 있는 행의 수가 전체 업체 수
                    valid_bidders = data_df[data_df.iloc[:, 0].notna()]
                    count = len(valid_bidders)
                    
                    # 파일명(또는 공고명으로 매칭하기 위해 파일명에 포함된 숫자로 식별할 수도 있으나,
                    # 시트1의 B열에 파일명과 유사한 '공고명'이 들어있을 것으로 예상)
                    # 여기서는 파일명에서 확장자를 제외한 부분으로 매핑해 보겠습니다.
                    # 하지만 시트1의 공고명 구조를 확인한 후에 매칭하는 것이 좋습니다.
                    # 우선 날짜(6자리) 기준으로 매핑을 시도하거나 이름으로 매핑합니다.
                    
                    # 공고명(파일명 기반 매칭 키)
                    key = file.replace('.xlsb', '').strip()
                    bidder_counts[key] = count
                    
                except Exception as e:
                    print(f"  -> 파일 처리 중 에러: {e}")
                    
    return bidder_counts

def add_col_to_sheet1():
    print("업체 수 데이터 추출 시작...")
    bidder_counts = extract_bidder_count(BASE_DIR)
    print(f"총 {len(bidder_counts)}건의 파일에서 업체 수 추출 완료.")
    
    print("\n구글 시트 연동 중...")
    client = get_google_sheet_client()
    sh = client.open_by_key(SHEET_ID)
    
    try:
        # gid=0 (첫 번째 시트)
        ws1 = sh.get_worksheet_by_id(0)
        
        all_data = ws1.get_all_values()
        if len(all_data) <= 1:
            print("데이터가 없습니다.")
            return

        # B열 오른쪽에 새로운 열 C 삽입
        ws1.insert_cols([['입찰업체 갯수']], 3) # index 3 means inserting AS column C (A=1, B=2, C=3)
        print("C열에 '입찰업체 갯수' 컬럼을 삽입했습니다. 매칭 시작...")
        
        # 다시 데이터를 읽어와서 매칭 진행 (C열이 생겼으므로 한 칸씩 밀림)
        all_data_new = ws1.get_all_values()
        
        updates = []
        # 헤더는 이미 C1에 '입찰업체 갯수'가 들어가 있거나 빈 칸일 수 있음
        updates.append({'range': 'C1', 'values': [['입찰업체 갯수']]})
        
        for r_idx in range(1, len(all_data_new)):
            row = all_data_new[r_idx]
            if len(row) > 1: # B열(index 1)이 존재하는지
                b_val = row[1].strip() # B열 텍스트 (공고명으로 추정)
                
                # B열의 공고명과 추출한 엑셀 파일명을 비교하여 매칭
                matched_count = ""
                for key, count in bidder_counts.items():
                    # 이름이 포함되어 있는지 느슨하게 비교
                    if b_val in key or key in b_val:
                        matched_count = count
                        break
                        
                if matched_count != "":
                    # row는 0-indexed이므로 A=0, B=1, C=2
                    updates.append({
                        'range': f'C{r_idx+1}',
                        'values': [[matched_count]]
                    })
                    
        if updates:
            print(f"총 {len(updates)-1}건의 매칭된 업체 수를 업데이트합니다...")
            ws1.batch_update(updates, value_input_option='USER_ENTERED')
            print("업데이트 완료!")
        else:
            print("매칭되는 데이터가 없습니다.")
            
    except Exception as e:
        print(f"시트1 포맷팅 에러: {e}")

if __name__ == "__main__":
    add_col_to_sheet1()
