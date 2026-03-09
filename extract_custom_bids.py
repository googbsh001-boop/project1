import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import re
from gspread_formatting import *

sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1n3WxFMxjS-mhHGE8I4dXi4Q2oJ3l4sq_OkCkBeJkbJI'

if len(sys.argv) < 3:
    print("Usage: python extract_custom_bids.py <COMPANY_KEYWORD> <SHEET_NAME>")
    sys.exit(1)

COMPANY_KEYWORD = sys.argv[1]
WORKSHEET_NAME = sys.argv[2]
BASE_DIR = r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_01.종심제,종평제"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_google_sheet_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def extract_bid_date(filename, df):
    try:
        if isinstance(df, dict):
            for sheet_name, sheet_df in df.items():
                for idx, row in sheet_df.head(50).iterrows():
                    row_strs = [str(x) for x in row if pd.notna(x)]
                    if any("개찰일" in s or "입찰일" in s or "투찰일" in s for s in row_strs):
                        for c_idx, val in enumerate(row):
                            if isinstance(val, str) and ("개찰일" in val or "입찰일" in val or "투찰일" in val):
                                if c_idx + 1 < len(row) and pd.notna(row[c_idx+1]):
                                    date_val = str(row[c_idx+1])
                                    match = re.search(r'(\d{4})[-/년]', date_val)
                                    if match: return match.group(1), date_val
        else:
            for idx, row in df.head(50).iterrows():
                row_strs = [str(x) for x in row if pd.notna(x)]
                if any("개찰일" in s or "입찰일" in s or "투찰일" in s for s in row_strs):
                    for c_idx, val in enumerate(row):
                        if isinstance(val, str) and ("개찰일" in val or "입찰일" in val or "투찰일" in val):
                            if c_idx + 1 < len(row) and pd.notna(row[c_idx+1]):
                                date_val = str(row[c_idx+1])
                                match = re.search(r'(\d{4})[-/년]', date_val)
                                if match: return match.group(1), date_val
    except Exception:
        pass

    match = re.search(r'(\d{2})(\d{2})(\d{2})', filename)
    if match:
        yy = int(match.group(1))
        mm = match.group(2)
        dd = match.group(3)
        year = 2000 + yy if yy < 50 else 1900 + yy
        return str(year), f"{year}-{mm}-{dd}"
        
    return "연도미상", "일자미상"

def extract_project_name(filename):
    name = filename
    if " - " in name:
        name = name.split(" - ", 1)[1]
    name = re.sub(r'^\d{6}\s*', '', name)
    name = re.sub(r'\(종심.*?\)\s*', '', name)
    name = re.sub(r'\(종평.*?\)\s*', '', name)
    name = re.sub(r'_Rev.*$', '', name)
    name = re.sub(r'\.[a-zA-Z]+$', '', name)
    name = name.replace("입찰결과", "").strip()
    if name.startswith("- "):
        name = name[2:]
    return name.strip()

def process_files():
    results = []
    excel_files = []
    
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if file.endswith(('.xls', '.xlsx', '.xlsb')) and not file.startswith('~'):
                if any(str(y) in root for y in range(2014, 2024)):
                    continue
                excel_files.append((root, file))
                
    print(f"총 {len(excel_files)}개의 엑셀 파일을 스캔합니다.")
    
    processed_count = 0
    for root, file in excel_files:
        fpath = os.path.join(root, file)
        project_name = extract_project_name(file)
        
        try:
            if file.endswith('.xlsb'):
                df_all = pd.read_excel(fpath, sheet_name=None, engine='pyxlsb', header=None)
                df = df_all[list(df_all.keys())[0]] 
            else:
                df_all = pd.read_excel(fpath, sheet_name=None, header=None)
                df = df_all[list(df_all.keys())[0]]
                
            year, full_date = extract_bid_date(file, df_all)
            
            try:
                if int(year) < 2024:
                    continue
            except:
                pass
                
            for idx, row in df.iterrows():
                row_strs = [str(x) for x in row if pd.notna(x)]
                if any(COMPANY_KEYWORD.lower() in s.replace(" ", "").lower() for s in row_strs):
                    n_val = row[13] if len(row) > 13 else None
                    rank_val = ""
                    if pd.notna(n_val):
                        try:
                            if float(n_val).is_integer():
                                rank_val = str(int(float(n_val)))
                            else:
                                rank_val = str(n_val)
                        except:
                            rank_val = str(n_val)
                            
                    bid_amount = row[2] if len(row) > 2 else ""
                    bid_ratio = row[4] if len(row) > 4 else ""
                    
                    try:
                        amt_str = str(bid_amount).replace(',', '').strip()
                        if amt_str == "" or amt_str.lower() == "nan":
                            continue
                        float(amt_str)
                    except ValueError:
                        continue
                        
                    results.append([
                        year,
                        full_date,
                        project_name,
                        "O",
                        rank_val,
                        bid_amount,
                        bid_ratio,
                        file
                    ])
                    # Found, skip rest
                    # 공사 하나 당 여러개 컨소시엄이면?
                    # 중복방지에서 제거됨
        except Exception as e:
            pass
            
        processed_count += 1
        if processed_count % 10 == 0:
            print(f"{processed_count}/{len(excel_files)} 처리중...")
            
    results.sort(key=lambda x: (str(x[0]), str(x[1])), reverse=False)
    
    dedup_results = []
    seen_projects = set()
    for row in results:
        project_name = row[2]
        if project_name not in seen_projects:
            seen_projects.add(project_name)
            dedup_results.append(row)
            
    return dedup_results

def main():
    print(f"{COMPANY_KEYWORD} 2024년 이후 참여 공사 데이터 추출 시작...")
    extracted_data = process_files()
    
    if not extracted_data:
        print(f"조건에 맞는 데이터({COMPANY_KEYWORD} 참여)를 찾지 못했습니다.")
        extracted_data = [] 
    else:
        print(f"총 {len(extracted_data)} 건의 {COMPANY_KEYWORD} 참여 공사 데이터를 찾았습니다.")
    
    print("구글 시트 연동 중...")
    client = get_google_sheet_client()
    sh = client.open_by_key(SHEET_ID)
    
    try:
        worksheet = sh.worksheet(WORKSHEET_NAME)
        worksheet.clear()
        print(f"기존 '{WORKSHEET_NAME}' 시트를 초기화했습니다.")
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sh.add_worksheet(title=WORKSHEET_NAME, rows="100", cols="10")
        print(f"새 시트 '{WORKSHEET_NAME}'를 생성했습니다.")
        
    headers = [["연도", "입찰일자", "공사명", f"{COMPANY_KEYWORD[:2]}참여", "낙찰우선순위", "입찰금액", "기초대비 투찰율(%)", "원본파일명"]]
    
    cleaned_data = []
    for row in extracted_data:
        cleaned_row = ["" if pd.isna(val) else val for val in row]
        cleaned_data.append(cleaned_row)
        
    final_data = headers + cleaned_data
    
    worksheet.update(values=final_data, range_name='A1')
    
    print("서식 적용 중...")
    fmt_header = CellFormat(
        backgroundColor=Color(0.85, 0.9, 0.95),
        textFormat=TextFormat(bold=True),
        horizontalAlignment='CENTER'
    )
    
    end_row = max(len(final_data), 2)
    fmt_currency = CellFormat(numberFormat=NumberFormat(type='NUMBER', pattern='#,##0"원"'))
    fmt_percent = CellFormat(numberFormat=NumberFormat(type='PERCENT', pattern='0.000%'))
    fmt_center = CellFormat(horizontalAlignment='CENTER')
    
    format_cell_ranges(worksheet, [
        ('A1:H1', fmt_header),
        (f'F2:F{end_row}', fmt_currency),
        (f'G2:G{end_row}', fmt_percent),
        (f'A2:B{end_row}', fmt_center),
        (f'D2:E{end_row}', fmt_center)
    ])
    
    rule = ConditionalFormatRule(
        ranges=[GridRange.from_a1_range(f'A2:H{end_row}', worksheet)],
        booleanRule=BooleanRule(
            condition=BooleanCondition('CUSTOM_FORMULA', ['=OR($E2="1", $E2=1)']),
            format=CellFormat(backgroundColor=Color(1.0, 0.95, 0.8))
        )
    )
    rules = get_conditional_format_rules(worksheet)
    rules.clear()
    rules.append(rule)
    rules.save()
    
    set_column_width(worksheet, 'A', 60)
    set_column_width(worksheet, 'B', 100)
    set_column_width(worksheet, 'C', 350)
    set_column_width(worksheet, 'D', 80)
    set_column_width(worksheet, 'E', 100)
    set_column_width(worksheet, 'F', 150)
    set_column_width(worksheet, 'G', 120)
    set_column_width(worksheet, 'H', 400)
    
    print("모든 작업이 완료되었습니다.")

if __name__ == "__main__":
    main()
