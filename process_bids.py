import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime

# --- 설정 ---
CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1ZwmfkDFJDBLK_U2oymeI3XZ3w5FLLak7AxodXN1pXrU'
WORKSHEET_NAME = '입찰결과정리' # 새로운 시트 이름
FOLDER_PATH = r"E:\인프라수주팀\입찰결과분석"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_google_sheet_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def process_file(file_path):
    print(f"Reading file: {os.path.basename(file_path)}")
    try:
        # 헤더 없이 읽어서 위치 기반으로 데이터 추출
        df = pd.read_excel(file_path, engine='pyxlsb', header=None)
        
        # 공사명 추출 (3행 2열 -> 인덱스[3][2])
        project_name = df.iloc[3, 2]
        print(f"Project Code/Name: {project_name}")

        # 데이터 테이블 시작 위치 찾기 ('순위'가 있는 행)
        header_row_idx = -1
        for idx, row in df.iterrows():
            # 0번 컬럼이 '순위'인 행 찾기
            if str(row[0]).replace(" ", "") == "순위":
                header_row_idx = idx
                break
        
        if header_row_idx == -1:
            print("Could not find data table header.")
            return None

        # 데이터 추출 (헤더 다음 행부터)
        data_df = df.iloc[header_row_idx+1:].copy()
        
        # 필요한 컬럼만 선택 (순위, 회사명, 입찰금액, 예가대비)
        # 인덱스: 0, 1, 2, 3
        summary_data = []
        for _, row in data_df.iterrows():
            rank = row[0]
            if pd.isna(rank): # 순위가 없으면 데이터 끝으로 간주
                continue
            
            try:
                # 데이터 정제
                company = row[1]
                amount = row[2]
                ratio = row[3]
                
                # NaN 처리
                if pd.isna(company): company = ""
                if pd.isna(amount): amount = ""
                if pd.isna(ratio): ratio = ""
                
                summary_data.append([rank, company, amount, ratio])
            except Exception as e:
                continue
                
        return project_name, summary_data

    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return None

def main():
    client = get_google_sheet_client()
    sh = client.open_by_key(SHEET_ID)
    
    # 워크시트 준비 (없으면 생성, 있으면 초기화)
    try:
        worksheet = sh.worksheet(WORKSHEET_NAME)
        worksheet.clear()
    except gspread.WorksheetNotFound:
        worksheet = sh.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=10)
    
    # 파일 목록 가져오기
    files = [f for f in os.listdir(FOLDER_PATH) if f.endswith('.xlsb')]
    files.sort() # 파일 이름 순 정렬

    all_rows = []
    
    # 헤더 추가
    all_rows.append(["입찰 결과 요약 정리", f"업데이트: {datetime.datetime.now()}"])
    all_rows.append([]) # 빈 줄

    for file_name in files:
        file_path = os.path.join(FOLDER_PATH, file_name)
        result = process_file(file_path)
        
        if result:
            project_name, data = result
            
            # 프로젝트 제목
            all_rows.append([f"■ 공사명: {project_name}"])
            # 테이블 헤더
            all_rows.append(["순위", "회사명", "입찰금액", "예가대비"])
            # 데이터
            all_rows.extend(data)
            # 구분선 (빈 줄)
            all_rows.append([])
            all_rows.append([])

    # 구글 시트에 쓰기
    if all_rows:
        print("Uploading to Google Sheet...")
        worksheet.update(all_rows)
        print("Done!")
    else:
        print("No data found to upload.")

if __name__ == "__main__":
    main()
