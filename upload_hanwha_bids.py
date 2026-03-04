import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# 인코딩 설정 (Windows Cmd 등에서 출력 오류 방지)
sys.stdout.reconfigure(encoding='utf-8')

# --- 설정 ---
CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1ZwmfkDFJDBLK_U2oymeI3XZ3w5FLLak7AxodXN1pXrU'
WORKSHEET_NAME = '시트2'
BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_google_sheet_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def extract_hanwha_data(base_dir):
    result = []
    
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.endswith('.xlsb') and not file.startswith('~'):
                file_path = os.path.join(root, file)
                print(f"엑셀 읽는 중: {file}")
                try:
                    df = pd.read_excel(file_path, engine='pyxlsb', header=None)
                    
                    # 헤더 행 찾기
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
                        # '기초대비' or '투찰율'
                        idx_ratio = -1
                        if '기초대비' in header_row:
                            idx_ratio = header_row.index('기초대비')
                        elif '투찰율' in header_row:
                            idx_ratio = header_row.index('투찰율')
                            
                        # '낙찰우선순위' or '우선순위' or '순위'
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

                    # 데이터 행에서 한화건설 찾기
                    data_df = df.iloc[header_row_idx+1:]
                    participated = "X"
                    ratio_val = ""
                    priority_val = ""
                    
                    for _, row in data_df.iterrows():
                        company = str(row[idx_company]).strip()
                        if "한화" in company:
                            participated = "O"
                            ratio = row[idx_ratio] if idx_ratio != -1 and not pd.isna(row[idx_ratio]) else ""
                            if ratio != "":
                                # 비율로 저장 (구글 시트에서 퍼센트 포맷을 적용할 수 있도록)
                                ratio_val = ratio
                            
                            priority = row[idx_priority] if idx_priority != -1 and not pd.isna(row[idx_priority]) else ""
                            if priority != "":
                                priority_val = priority
                            break
                            
                    result.append([file, participated, ratio_val, priority_val])
                    
                except Exception as e:
                    print(f"  -> 파일 처리 중 에러: {e}")
                    
    return result

def main():
    print("구글 시트 인증 시도...")
    client = get_google_sheet_client()
    sh = client.open_by_key(SHEET_ID)
    
    # 시트2 가져오기 (없으면 생성)
    try:
        worksheet = sh.worksheet(WORKSHEET_NAME)
        worksheet.clear() # 기존 데이터 초기화
        print(f"'{WORKSHEET_NAME}' 시트를 찾았습니다. 데이터를 초기화합니다.")
    except gspread.WorksheetNotFound:
        worksheet = sh.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=20)
        print(f"'{WORKSHEET_NAME}' 시트를 새로 생성했습니다.")
        
    print("\n로컬 엑셀 파일 스캔 시작...")
    extracted_data = extract_hanwha_data(BASE_DIR)
    
    # 데이터 정리
    headers = [["파일명", "입찰유무(한화)", "기초대비 투찰율(%)", "낙찰우선순위"]]
    
    final_data = headers + extracted_data
    
    print("\n구글 시트에 데이터 업로드 중...")
    worksheet.update('A1', final_data)
    
    print("\n작업 완료! 구글 시트를 확인해 주세요.")

if __name__ == "__main__":
    main()
