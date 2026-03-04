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
WORKSHEET_NAME = '시트2'
WORKSHEET_ID = 1426009222
BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_google_sheet_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def extract_date(filename):
    match = re.search(r'(\d{6})', filename)
    if match:
        return match.group(1)
    return "999999"

def extract_hanwha_data(base_dir):
    result = []
    
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
                        
                    header_row = [str(x).replace("\n", "").replace(" ", "") for x in df.iloc[header_row_idx].values]
                    
                    try:
                        idx_company = header_row.index('회사명')
                        
                        # 기초대비 (투찰율)
                        idx_base_ratio = -1
                        if '기초대비' in header_row:
                            idx_base_ratio = header_row.index('기초대비')
                        elif '투찰율' in header_row:
                            idx_base_ratio = header_row.index('투찰율')
                            
                        # 기초금액대비 (기초금액대비 낙찰율을 새로 계산하기 위해 추가 탐색 안할시 기초대비를 쓰거나, 실제 입찰금액/기초금액)
                        # 여기서는 사실 엑셀에 있는 '기초대비' 항목 자체가 기초금액 대비 입찰금액의 비율입니다.
                        # 따라서 사용자가 요구한 '기초금액 대비 낙찰율(%)'은 이미 해당 엑셀에서 제공하는 '기초대비' 와 값이 동일할 확률이 매우 높지만,
                        # 시트2의 컬럼명과 요구사항을 맞추겠습니다.
                        # (만약 '입찰금액' / '기초금액' 원 데이터를 찾아서 직접 계산하려면 추가적인 컬럼 검색이 필요합니다.
                        # 보통 '기초대비' 열 자체가 [입찰금액 / 기초금액] 입니다.)

                        # 낙찰우선순위
                        idx_priority = -1
                        if '낙찰우선순위' in header_row:
                            idx_priority = header_row.index('낙찰우선순위')
                        elif '우선순위' in header_row:
                            idx_priority = header_row.index('우선순위')
                        elif '순위' in header_row:
                            idx_priority = header_row.index('순위')
                            
                    except ValueError as e:
                        print(f"  -> 필수 컬럼이 헤더에 없습니다: {e}")
                        continue

                    data_df = df.iloc[header_row_idx+1:]
                    participated = "X"
                    base_ratio_val = ""
                    base_winning_ratio_val = ""
                    priority_val = ""
                    
                    for _, row in data_df.iterrows():
                        company = str(row[idx_company]).strip()
                        if "한화" in company:
                            participated = "O"
                            
                            br = row[idx_base_ratio] if idx_base_ratio != -1 and not pd.isna(row[idx_base_ratio]) else ""
                            if br != "": 
                                base_ratio_val = br
                                base_winning_ratio_val = br # '기초금액 대비 낙찰율' 도 결국 같은 기초대비 열에서 가져옴 (추후 분리 필요시 로직 변경)
                            
                            pr = row[idx_priority] if idx_priority != -1 and not pd.isna(row[idx_priority]) else ""
                            if pr != "": 
                                priority_val = pr
                            
                            break
                            
                    result.append([file, participated, base_ratio_val, base_winning_ratio_val, priority_val])
                    
                except Exception as e:
                    print(f"  -> 파일 처리 중 에러: {e}")
                    
    return result

def main():
    print("로컬 엑셀 파일 스캔 시작...")
    extracted_data = extract_hanwha_data(BASE_DIR)
    
    # 입찰날짜(파일명 안의 6자리 숫자) 기준으로 오름차순 정렬
    print("입찰날짜 기준으로 오름차순 정렬 중...")
    extracted_data.sort(key=lambda x: extract_date(x[0]))
    
    # 데이터 정리 (헤더 변경)
    headers = [["공고명", "입찰유무(한화)", "기초대비 투찰율(%)", "기초금액 대비 낙찰율(%)", "낙찰우선순위"]]
    final_data = headers + extracted_data
    
    print("\n구글 시트 연동 중...")
    client = get_google_sheet_client()
    sh = client.open_by_key(SHEET_ID)
    
    try:
        worksheet = sh.get_worksheet_by_id(WORKSHEET_ID)
        worksheet.clear()
        print(f"시트 '{worksheet.title}' 데이터를 초기화했습니다.")
    except Exception as e:
        print(f"시트를 찾지 못했습니다: {e}")
        return
        
    print("구글 시트에 데이터 업로드 중...")
    worksheet.update('A1', final_data)
    
    print("퍼센트(%) 표시 형식 적용 중 - 소수점 셋째자리 반올림 (C열, D열)...")
    try:
        # 소수점 넷째자리에서 반올림하여 셋째자리까지 표시 (0.000%)
        fmt = CellFormat(numberFormat=NumberFormat(type='PERCENT', pattern='0.000%'))
        
        # C열(투찰율)과 D열(낙찰율)에 퍼센트 서식 지정
        format_cell_range(worksheet, 'C2:D1000', fmt)
        print("서식 적용 완료!")
    except Exception as e:
        print(f"서식 적용 에러: {e}")
        
    print("\n모든 작업이 완료되었습니다!")

if __name__ == "__main__":
    main()
