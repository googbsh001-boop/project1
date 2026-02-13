import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import re
import openpyxl
from gspread_formatting import format_cell_ranges, CellFormat, Color
from gspread.utils import rowcol_to_a1

# --- 설정 ---
CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1ZwmfkDFJDBLK_U2oymeI3XZ3w5FLLak7AxodXN1pXrU'
WORKSHEET_NAME = '입찰결과분석_상세' 
FOLDER_PATH = r"E:\인프라수주팀\입찰결과분석"
COLOR_FILE_NAME = "계양~강화 고속국도 3~6공구 참여업체_조모임 포함.xlsx"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_google_sheet_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def load_company_colors():
    """조모임 파일에서 업체별 색상 그룹을 추출하여 google color 객체로 매핑"""
    file_path = os.path.join(FOLDER_PATH, COLOR_FILE_NAME)
    if not os.path.exists(file_path):
        print(f"Color file not found: {file_path}")
        return {}

    print(f"Loading colors from: {COLOR_FILE_NAME}")
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        
        company_colors = {}
        
        # 색상 매핑 (Theme/Tint -> Google Color)
        # 1. Theme 9 (경남 등) -> Light Green
        color_group_1 = Color(0.8, 1.0, 0.8) 
        # 2. Theme 6 (계룡 등) -> Light Orange/Pink
        color_group_2 = Color(1.0, 0.9, 0.8)
        # 3. Theme 8 (금호 등) -> Light Yellow/Gold
        color_group_3 = Color(1.0, 1.0, 0.8)
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    fill = cell.fill
                    if fill and fill.start_color:
                        color = fill.start_color
                        key = ""
                        if color.type == 'theme':
                            key = f"Theme_{color.theme}"
                        
                        target_color = None
                        if "Theme_9" in key: target_color = color_group_1
                        elif "Theme_6" in key: target_color = color_group_2
                        elif "Theme_8" in key: target_color = color_group_3
                        
                        if target_color:
                            company = cell.value.strip()
                            company_colors[company] = target_color
                            
        return company_colors
    except Exception as e:
        print(f"Error loading colors: {e}")
        return {}


def extract_zone(filename):
    match = re.search(r'제(\d+)공구', filename)
    if match:
        return f"{match.group(1)}공구"
    return "기타"

def process_file(file_path):
    print(f"Reading file: {os.path.basename(file_path)}")
    try:
        df = pd.read_excel(file_path, engine='pyxlsb', header=None)
        
        # 데이터 테이블 시작 위치 찾기
        header_row_idx = -1
        for idx, row in df.iterrows():
            if str(row[0]).replace(" ", "") == "순위":
                header_row_idx = idx
                break
        
        if header_row_idx == -1: return []

        data_rows = []
        data_df = df.iloc[header_row_idx+1:].copy()
        
        for _, row in data_df.iterrows():
            rank = row[0]
            if pd.isna(rank): continue
            
            rank_str = str(rank).strip()
            if not rank_str.isdigit(): continue

            try:
                rank_val = int(rank_str)
                company = str(row[1]).strip()
                # 1순위(낙찰사)에 별표 표시
                if rank_val == 1:
                    company = "★ " + company

                amount = float(row[2])
                ratio = float(row[4]) 
                
                amount_billions = amount / 100000000
                ratio_percent = ratio * 100
                
                data_rows.append({
                    'rank': rank_val,
                    'company': company,
                    'amount': amount_billions,
                    'ratio': ratio_percent
                })
            except (ValueError, TypeError):
                continue
        
        # 순위 기준으로 정렬
        data_rows.sort(key=lambda x: x['rank'])
        return data_rows

    except Exception as e:
        print(f"Error checking file {file_path}: {e}")
        return []

def main():
    client = get_google_sheet_client()
    sh = client.open_by_key(SHEET_ID)
    
    try:
        worksheet = sh.worksheet(WORKSHEET_NAME)
        worksheet.clear()
    except gspread.WorksheetNotFound:
        worksheet = sh.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=20)
    
    # 1. 파일별 데이터 수집
    files = [f for f in os.listdir(FOLDER_PATH) if f.endswith('.xlsb')]
    
    zone_data = {} # {'3공구': [row1, row2...], '4공구': ...}
    
    for file_name in files:
        zone = extract_zone(file_name)
        file_path = os.path.join(FOLDER_PATH, file_name)
        rows = process_file(file_path)
        zone_data[zone] = rows

    # 2. Side-by-Side 레이아웃 구성
    sorted_zones = sorted(zone_data.keys())
    
    upload_rows = []
    
    # 헤더 1: 공구 이름
    header1 = []
    # 헤더 2: 컬럼명
    header2 = []
    
    for zone in sorted_zones:
        # 공구명 + 빈칸 3개 (4열 차지)
        header1.extend([f"■ {zone}", "", "", ""]) 
        header2.extend(["순위", "회사명", "입찰금액(억원)", "기초대비(%)"])
    
    upload_rows.append(header1)
    upload_rows.append(header2)
    
    # 데이터 채우기
    # 가장 긴 데이터 길이 찾기
    max_len = 0
    for zone in sorted_zones:
        max_len = max(max_len, len(zone_data[zone]))
        
    for i in range(max_len):
        row_data = []
        for zone in sorted_zones:
            rows = zone_data[zone]
            if i < len(rows):
                entry = rows[i]
                row_data.extend([
                    entry['rank'],
                    entry['company'],
                    round(entry['amount'], 2),
                    round(entry['ratio'], 4)
                ])
            else:
                row_data.extend(["", "", "", ""]) # 데이터가 없으면 빈칸
        upload_rows.append(row_data)

    if upload_rows:
        print("Uploading data...")
        worksheet.update(upload_rows)
        
        # --- 스타일링 (하늘색 배경 적용) ---
        from gspread_formatting import format_cell_ranges, CellFormat, Color
        from gspread.utils import rowcol_to_a1

        print("Applying styles...")
        
        # 하늘색 (Light Blue) - 공구/순위 열 배경
        light_blue = Color(0.85, 0.93, 1.0) 
        fmt_blue = CellFormat(backgroundColor=light_blue)
        
        batch = []
        total_rows = len(upload_rows)
        
        # 1. 공구/순위 열 하늘색 적용
        for i in range(len(sorted_zones)):
            col_idx = i * 4 + 1
            start_cell = rowcol_to_a1(1, col_idx)
            end_cell = rowcol_to_a1(total_rows, col_idx)
            range_str = f"{start_cell}:{end_cell}"
            batch.append((range_str, fmt_blue))
            
        # 2. 업체별 색상 적용
        company_colors = load_company_colors()
        if company_colors:
            # 업로드된 데이터 순회하며 회사명 찾기 (3행부터 데이터 시작)
            # header1(1행), header2(2행) 제외
            for row_idx in range(2, total_rows): 
                row_data = upload_rows[row_idx]
                # 각 공구별 회사명 컬럼 위치: 1, 5, 9, 13 (0-based index applied to upload_rows list)
                # Google Sheet column index (1-based): 2, 6, 10, 14
                
                for i in range(len(sorted_zones)):
                    sheet_col_idx = i * 4 + 2 # 회사명 컬럼 (1-based)
                    list_col_idx = i * 4 + 1  # upload_rows 리스트 인덱스 (0-based)
                    
                    if list_col_idx < len(row_data):
                        company_val = str(row_data[list_col_idx]).replace("★ ", "").strip()
                        if company_val in company_colors:
                            cell_a1 = rowcol_to_a1(row_idx + 1, sheet_col_idx) # row_idx는 0-based, sheet는 1-based
                            fmt_company = CellFormat(backgroundColor=company_colors[company_val])
                            batch.append((cell_a1, fmt_company))

        format_cell_ranges(worksheet, batch)
        print("Done!")
    else:
        print("No data collected.")

if __name__ == "__main__":
    main()
