import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import re
import openpyxl
from gspread_formatting import format_cell_ranges, CellFormat, Color, TextFormat
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

# 색상 및 그룹 정의 (전역 변수로 관리)
GROUP_COLORS = {
    'Theme_9': {'name': 'Group A (Green)', 'color': Color(0.7, 0.9, 0.7)},
    'Theme_6': {'name': 'Group B (Orange)', 'color': Color(1.0, 0.8, 0.6)},
    'Theme_8': {'name': 'Group C (Yellow)', 'color': Color(1.0, 1.0, 0.6)}
}

def get_google_sheet_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def load_company_map():
    """조모임 파일에서 업체별 색상 및 그룹 정보를 추출"""
    file_path = os.path.join(FOLDER_PATH, COLOR_FILE_NAME)
    if not os.path.exists(file_path):
        print(f"Color file not found: {file_path}")
        return {}

    print(f"Loading colors from: {COLOR_FILE_NAME}")
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        
        company_map = {} # company -> {'group': 'Theme_X', 'color': ColorObj}
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    fill = cell.fill
                    if fill and fill.start_color:
                        color = fill.start_color
                        theme_key = ""
                        if color.type == 'theme':
                            theme_key = f"Theme_{color.theme}"
                        
                        matched_group_key = None
                        if "Theme_9" in theme_key: matched_group_key = 'Theme_9'
                        elif "Theme_6" in theme_key: matched_group_key = 'Theme_6'
                        elif "Theme_8" in theme_key: matched_group_key = 'Theme_8'
                        
                        if matched_group_key:
                            company = cell.value.strip()
                            company_map[company] = {
                                'group_key': matched_group_key,
                                'color': GROUP_COLORS[matched_group_key]['color']
                            }
                            
        return company_map
    except Exception as e:
        print(f"Error loading colors: {e}")
        return {}

def find_company_info(company_name, company_map):
    """회사명에 맞는 정보(색상, 그룹)를 부분 일치로 검색"""
    if not company_name: return None

    # 1. 완전 일치
    if company_name in company_map:
        return company_map[company_name]
    
    # 2. 정제 후 부분 일치
    clean_target = company_name.replace("주식회사", "").replace("(주)", "").replace(" ", "")
    
    if not clean_target: return None

    for key_name, info in company_map.items():
        clean_key = key_name.replace("주식회사", "").replace("(주)", "").replace(" ", "")
        
        if not clean_key: continue

        if len(clean_key) > 1 and (clean_key in clean_target or clean_target in clean_key):
             return info
             
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
    
    # 1. Excel 파일 처리
    files = [f for f in os.listdir(FOLDER_PATH) if f.endswith('.xlsb')]
    
    zone_data = {}
    for file_name in files:
        zone = extract_zone(file_name)
        file_path = os.path.join(FOLDER_PATH, file_name)
        rows = process_file(file_path)
        zone_data[zone] = rows

    sorted_zones = sorted(zone_data.keys())
    company_map = load_company_map() # 색상 정보 로드
    
    # 2. 메인 데이터 구성
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

    # 3. 평균 계산 (Groups per Zone)
    summary_start_row = len(upload_rows) + 3 # 메인 데이터와 요약 표 사이 간격
    summary_rows = []
    
    # 요약 테이블 헤더
    summary_header = ["구분"]
    for zone in sorted_zones:
        summary_header.append(f"{zone} 평균(%)")
    summary_rows.append(summary_header)
    
    # 그룹별 로우 생성
    groups = ['Theme_9', 'Theme_6', 'Theme_8'] # 순서 지정
    group_display_names = {
        'Theme_9': "Color Group 1 (Green)",
        'Theme_6': "Color Group 2 (Orange)",
        'Theme_8': "Color Group 3 (Yellow)"
    }
    
    for group_key in groups:
        row = [group_display_names[group_key]]
        for zone in sorted_zones:
            # 해당 공구, 해당 그룹의 ratio 수집
            rows = zone_data[zone]
            ratios = []
            for entry in rows:
                clean_name = str(entry['company']).replace("★ ", "").strip()
                info = find_company_info(clean_name, company_map)
                if info and info['group_key'] == group_key:
                    ratios.append(entry['ratio'])
            
            if ratios:
                avg = sum(ratios) / len(ratios)
                row.append(round(avg, 4))
            else:
                row.append("-")
        summary_rows.append(row)

    # 4. 데이터 업로드 (메인 + 요약)
    final_payload = upload_rows + [[] for _ in range(2)] + summary_rows # 2줄 빈칸 추가
    
    if final_payload:
        print("Uploading data...")
        worksheet.update(final_payload)
        
        # --- 스타일링 ---
        print("Applying styles...")
        
        batch = []
        total_rows = len(upload_rows) # 메인 데이터만 색칠하기 위해
        
        # [메인] 하늘색 (공구/순위 열)
        light_blue = Color(0.85, 0.93, 1.0) 
        fmt_blue = CellFormat(backgroundColor=light_blue)
        
        for i in range(len(sorted_zones)):
            col_idx = i * 4 + 1
            start_cell = rowcol_to_a1(1, col_idx)
            end_cell = rowcol_to_a1(total_rows, col_idx) # 메인 데이터 높이까지만
            range_str = f"{start_cell}:{end_cell}"
            batch.append((range_str, fmt_blue))
            
        # [메인] 업체별 색상 적용 (부분 일치)
        if company_map:
            for row_idx in range(2, total_rows): 
                row_data = upload_rows[row_idx]
                for i in range(len(sorted_zones)):
                    sheet_col_idx = i * 4 + 2 
                    list_col_idx = i * 4 + 1  
                    if list_col_idx < len(row_data):
                        original_val = str(row_data[list_col_idx])
                        company_val = original_val.replace("★ ", "").strip()
                        if not company_val: continue
                        
                        info = find_company_info(company_val, company_map)
                        if info:
                            cell_a1 = rowcol_to_a1(row_idx + 1, sheet_col_idx) 
                            fmt_company = CellFormat(backgroundColor=info['color'])
                            batch.append((cell_a1, fmt_company))

        # [요약] 테이블 스타일링
        # summary_start_row (1-based index calculation required)
        # final_payload 상에서 summary title은: len(upload_rows) + 2 (빈줄 2개 후) index -> +1 for sheet row
        sheet_summary_start_row = len(upload_rows) + 3
        
        # 요약 헤더 볼드체
        header_range = f"A{sheet_summary_start_row}:E{sheet_summary_start_row}" # 4개 공구 + 구분 = 5열(E)
        fmt_bold = CellFormat(textFormat=TextFormat(bold=True))
        batch.append((header_range, fmt_bold))
        
        # 요약 행 배경색 적용
        for idx, group_key in enumerate(groups):
            row_num = sheet_summary_start_row + 1 + idx
            # A열(이름)에 색상 적용
            cell_addr = f"A{row_num}"
            color = GROUP_COLORS[group_key]['color']
            
            # 셀 배경색
            fmt_group = CellFormat(backgroundColor=color, textFormat=TextFormat(bold=True))
            batch.append((cell_addr, fmt_group))

        print(f"Applying {len(batch)} format changes...")
        format_cell_ranges(worksheet, batch)
        print("Done!")
    else:
        print("No data collected.")

if __name__ == "__main__":
    main()
