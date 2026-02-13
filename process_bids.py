import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import re
import openpyxl
from gspread_formatting import format_cell_ranges, CellFormat, Color
from gspread.utils import rowcol_to_a1

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

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
        # 1. Theme 9 (경남 등) -> Green (더 진한 쑥색/녹색)
        color_group_1 = Color(0.7, 0.9, 0.7) 
        
        # 2. Theme 6 (계룡 등) -> Orange (더 진한 살구색/주황)
        color_group_2 = Color(1.0, 0.8, 0.6)
        
        # 3. Theme 8 (금호 등) -> Yellow (진한 노랑)
        color_group_3 = Color(1.0, 1.0, 0.6)

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

def find_color_for_company(company_name, color_map):
    """회사명에 맞는 색상을 부분 일치로 검색"""
    if not company_name: return None

    # 1. 완전 일치
    if company_name in color_map:
        return color_map[company_name]
    
    # 2. 정제 후 부분 일치
    # (주), 주식회사, 공백 제거
    clean_target = company_name.replace("주식회사", "").replace("(주)", "").replace(" ", "")
    
    if not clean_target: return None

    for key_name, color in color_map.items():
        clean_key = key_name.replace("주식회사", "").replace("(주)", "").replace(" ", "")
        
        if not clean_key: continue

        # 키가 타겟에 포함되거나, 타겟이 키에 포함되는 경우
        if len(clean_key) > 1 and (clean_key in clean_target or clean_target in clean_key):
             return color
             
    return None

def extract_zone(filename):
    match = re.search(r'제(\d+)공구', filename)
    if match:
        return f"{match.group(1)}공구"
    return "기타"

def process_file(file_path):
    print(f"Reading file: {os.path.basename(file_path)}")
    try:
        df = pd.read_excel(file_path, engine='pyxlsb', header=None)
        
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
    
    files = [f for f in os.listdir(FOLDER_PATH) if f.endswith('.xlsb')]
    
    zone_data = {}
    for file_name in files:
        zone = extract_zone(file_name)
        file_path = os.path.join(FOLDER_PATH, file_name)
        rows = process_file(file_path)
        zone_data[zone] = rows

    sorted_zones = sorted(zone_data.keys())
    
    upload_rows = []
    
    header1 = []
    header2 = []
    
    for zone in sorted_zones:
        header1.extend([f"■ {zone}", "", "", ""]) 
        header2.extend(["순위", "회사명", "입찰금액(억원)", "기초대비(%)"])
    
    upload_rows.append(header1)
    upload_rows.append(header2)
    
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
                row_data.extend(["", "", "", ""])
        upload_rows.append(row_data)

    if upload_rows:
        print("Uploading data...")
        worksheet.update(upload_rows)
        
        # --- 스타일링 ---
        print("Applying styles...")
        
        # 하늘색 (공구/순위 열)
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
            
        # 2. 업체별 색상 적용 (부분 일치)
        company_colors = load_company_colors()
        color_batch_size = 0
        
        if company_colors:
            print(f"Loaded {len(company_colors)} company colors for matching.")
            
            for row_idx in range(2, total_rows): 
                row_data = upload_rows[row_idx]
                
                for i in range(len(sorted_zones)):
                    sheet_col_idx = i * 4 + 2 
                    list_col_idx = i * 4 + 1  
                    
                    if list_col_idx < len(row_data):
                        original_val = str(row_data[list_col_idx])
                        company_val = original_val.replace("★ ", "").strip()
                        
                        if not company_val: continue
                        
                        # 부분 일치 검색
                        matched_color = find_color_for_company(company_val, company_colors)
                        
                        if matched_color:
                            cell_a1 = rowcol_to_a1(row_idx + 1, sheet_col_idx) 
                            fmt_company = CellFormat(backgroundColor=matched_color)
                            batch.append((cell_a1, fmt_company))
                            color_batch_size += 1
                        elif row_idx < 10: 
                             print(f"No match for: '{company_val}'")

        print(f"Applying {len(batch)} format changes...")
        format_cell_ranges(worksheet, batch)
        print("Done!")
    else:
        print("No data collected.")

if __name__ == "__main__":
    main()
