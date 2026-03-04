import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
from gspread_formatting import *
from gspread.utils import rowcol_to_a1

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
                    priority_val = ""
                    
                    # 1등 업체의 "E열 기초대비" (E열 = index 4, 단 E열인지 헤더의 기초대비 컬럼인지)
                    # 사용자가 명시적으로 "e열 기초대비"라고 했고, "n열의 우선순위" (N열 = index 13)라고 했습니다.
                    # N열(index 13)이 1인 로우의 E열(index 4) 값을 찾습니다.
                    est_winning_ratio_val = ""
                    
                    # 먼저 1위 낙찰율 찾기 (N열 우선순위 1)
                    # 파일에 따라 N열이 없을 수 있으므로 예외 처리를 합니다.
                    for _, row in data_df.iterrows():
                        try:
                            # N열 = 13
                            if len(row) > 13:
                                n_val = str(row[13]).strip()
                                # E열 = 4
                                if n_val == '1' and len(row) > 4:
                                    e_val = row[4]
                                    if not pd.isna(e_val):
                                        est_winning_ratio_val = e_val
                                        break
                        except Exception:
                            pass
                            
                    # 한화건설 본인 데이터 찾기
                    for _, row in data_df.iterrows():
                        company = str(row[idx_company]).strip()
                        if "한화" in company:
                            participated = "O"
                            
                            br = row[idx_base_ratio] if idx_base_ratio != -1 and not pd.isna(row[idx_base_ratio]) else ""
                            if br != "": 
                                base_ratio_val = br
                            
                            pr = row[idx_priority] if idx_priority != -1 and not pd.isna(row[idx_priority]) else ""
                            if pr != "": 
                                priority_val = pr
                            break
                            
                    result.append([file, participated, base_ratio_val, est_winning_ratio_val, priority_val])
                    
                except Exception as e:
                    print(f"  -> 파일 처리 중 에러: {e}")
                    
    return result

def main():
    print("로컬 엑셀 파일 스캔 시작...")
    extracted_data = extract_hanwha_data(BASE_DIR)
    
    print("입찰날짜 기준으로 오름차순 정렬 중...")
    extracted_data.sort(key=lambda x: extract_date(x[0]))
    
    headers = [["공고명", "입찰유무(한화)", "한화 기초대비 투찰율(%)", "기초금액 대비 낙찰율(%)", "낙찰우선순위"]]
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
    
    print("서식 적용 시작...")
    try:
        # 서식을 일괄 적용하기 위해 batch 리스트 준비
        batch = []
        
        # 1. 퍼센트 포맷 (C, D열 소수점 셋째자리)
        fmt_percent = CellFormat(numberFormat=NumberFormat(type='PERCENT', pattern='0.000%'))
        batch.append(('C2:D1000', fmt_percent))
        
        # 2. 낙찰우선순위(E열)가 10 미만인 행 배경색 적용
        # E열이 헤더 포함 5번째 열. 데이터는 행 2번부터.
        light_green = Color(0.85, 0.95, 0.85) # 보기 좋은 연녹색
        fmt_highlight = CellFormat(backgroundColor=light_green, textFormat=TextFormat(bold=True))
        
        # 행 개수 파악 후 개별 행 포맷팅 룰 추가
        total_rows = len(final_data)
        for data_idx, row_data in enumerate(extracted_data):
            row_num = data_idx + 2 # 헤더가 1번
            priority_val = row_data[4] # 낙찰우선순위
            
            # 우선순위가 유효한 숫자인지 확인
            if priority_val != "":
                try:
                    priority_int = int(float(priority_val))
                    if priority_int < 10:
                        # 해당 행 전체(A열~E열) 색상 칠하기
                        cell_range = f'A{row_num}:E{row_num}'
                        batch.append((cell_range, fmt_highlight))
                except ValueError:
                    # 숫자가 아닐 경우 넘어감
                    pass
        
        print(f"{len(batch)}건의 포맷 변경을 적용합니다...")
        format_cell_ranges(worksheet, batch)
        print("서식 적용 및 행 하이라이트 완료!")
        
    except Exception as e:
        print(f"서식 적용 에러: {e}")
        
    print("\n모든 작업이 완료되었습니다!")

if __name__ == "__main__":
    main()
