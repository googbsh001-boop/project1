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
WORKSHEET_NAME = '진흥기업 참여현황'
BASE_DIR = r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_google_sheet_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def extract_bid_date(filename, df):
    # 1. 시트 내에서 "입찰일", "개찰일", "투찰일" 검색
    try:
        if isinstance(df, dict): # xlsb 등 sheet_name=None 인 경우
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

    # 2. 파일명에서 날짜(YYMMDD) 추출 (Fall-back)
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
    return name.strip()

def process_files():
    results = []
    excel_files = []
    
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if file.endswith(('.xls', '.xlsx', '.xlsb')) and not file.startswith('~'):
                # 24년~26년 데이터 필터링 조건 추가 (폴더명이나 파일명 기준 1차 필터)
                # 정밀 필터링은 date 추출 후 수행하기 위해 여기서는 최대한 포함하거나 명시적으로 과거 제외
                if "2020전" in root or "2021" in root or "2022" in root or "2023" in root:
                    continue
                excel_files.append((root, file))
                
    print(f"총 {len(excel_files)}개의 엑셀 파일을 스캔합니다.")
    
    for root, file in excel_files:
        fpath = os.path.join(root, file)
        project_name = extract_project_name(file)
        
        try:
            if file.endswith('.xlsb'):
                df_all = pd.read_excel(fpath, sheet_name=None, engine='pyxlsb', header=None)
                df = df_all[list(df_all.keys())[0]] # 첫번째 시트를 메인 데이터로 사용
            else:
                df_all = pd.read_excel(fpath, sheet_name=None, header=None)
                df = df_all[list(df_all.keys())[0]]
                
            year, full_date = extract_bid_date(file, df_all)
            
            # 24~26년 필터링 (최종)
            try:
                if int(year) < 2024 or int(year) > 2026:
                    continue
            except:
                pass # 연도미상인 경우 혹시 모르니 일단 포함 (또는 제외 결정 필요)
                
            for idx, row in df.iterrows():
                row_strs = [str(x) for x in row if pd.notna(x)]
                if any("진흥" in s for s in row_strs):
                    n_val = row[13] if len(row) > 13 else None
                    rank_val = ""
                    if pd.notna(n_val):
                        # Convert float like 1.0 to integer string 1 if possible
                        try:
                            if float(n_val).is_integer():
                                rank_val = str(int(float(n_val)))
                            else:
                                rank_val = str(n_val)
                        except:
                            rank_val = str(n_val)
                            
                    # 사용자 피드백 반영: 투찰금액 = 입찰금액 (C열, 인덱스 2) / 투찰율 = 기초대비(%) (E열, 인덱스 4)
                    bid_amount = row[2] if len(row) > 2 else ""
                    bid_ratio = row[4] if len(row) > 4 else ""
                    
                    # E, F열이 누락되는 가짜 행(메타데이터/요약 행) 필터링
                    # 실제 입찰 데이터라면 C열(투찰금액)이 숫자여야 함
                    try:
                        amt_str = str(bid_amount).replace(',', '').strip()
                        if amt_str == "" or amt_str.lower() == "nan":
                            continue
                        float(amt_str)
                    except ValueError:
                        # 숫자로 변환할 수 없는 텍스트(예: "참여회사: 진흥기업 등")가 들어간 행은 실제 입찰 행이 아니므로 제외
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
        except Exception as e:
            print(f"Error reading {file}: {e}")
            
    # 연도별, 입찰일자별 오름차순 정렬 (과거 -> 최신)
    results.sort(key=lambda x: (str(x[0]), str(x[1])), reverse=False)
    
    # 공사명(C열, idx=2) 기준 중복 제거 (정렬 후 처음 나온 건만 유지)
    dedup_results = []
    seen_projects = set()
    for row in results:
        project_name = row[2]
        if project_name not in seen_projects:
            seen_projects.add(project_name)
            dedup_results.append(row)
            
    return dedup_results

def main():
    print("진흥기업 모든 참여 공사 데이터 추출 시작...")
    extracted_data = process_files()
    
    if not extracted_data:
        print("조건에 맞는 데이터(진흥기업 참여)를 찾지 못했습니다.")
        # 빈 시트라도 만들기 위해 계속 진행
        extracted_data = []
    else:
        print(f"총 {len(extracted_data)} 건의 진흥기업 참여 공사 데이터를 찾았습니다.")
    
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
        
    headers = [["연도", "입찰일자", "공사명", "진흥참여", "N열순위", "투찰금액", "투찰율", "원본파일명"]]
    
    # Replace any NaN or NAType values with empty string to avoid JSON errors
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
    # Check if we have data to style
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
    
    # 1순위 조건부 서식(색채우기) 추가
    rule = ConditionalFormatRule(
        ranges=[GridRange.from_a1_range(f'A2:H{end_row}', worksheet)],
        booleanRule=BooleanRule(
            condition=BooleanCondition('CUSTOM_FORMULA', ['=OR($E2="1", $E2=1)']),
            format=CellFormat(backgroundColor=Color(1.0, 0.95, 0.8)) # 연한 주황/노랑색으로 하이라이트
        )
    )
    rules = get_conditional_format_rules(worksheet)
    rules.clear() # 기존 룰 지우기
    rules.append(rule)
    rules.save()
    
    set_column_width(worksheet, 'A', 60)
    set_column_width(worksheet, 'B', 100)
    set_column_width(worksheet, 'C', 300)
    set_column_width(worksheet, 'D', 80)
    set_column_width(worksheet, 'E', 80)
    set_column_width(worksheet, 'F', 150)
    set_column_width(worksheet, 'G', 80)
    set_column_width(worksheet, 'H', 400)
    
    print("모든 작업이 완료되었습니다.")

if __name__ == "__main__":
    main()
