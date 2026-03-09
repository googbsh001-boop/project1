import os
import sys
import pandas as pd
import re

sys.stdout.reconfigure(encoding='utf-8')
base_path = r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사"

excel_files = []
for root, dirs, files in os.walk(base_path):
    for file in files:
        if file.endswith(('.xls', '.xlsx', '.xlsb')) and not file.startswith('~'):
            excel_files.append(os.path.join(root, file))

# Randomly select a few files from different years
test_files = [
    r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사\2020전\170103 (종심-일)대청댐계통 광역상수도사업 제2공구 정수시설공사.xlsx",
    r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사\2021\210430 [한전] 광주지역 전기공급시설 전력구공사(계림-일곡)2차.xls",
    r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사\2022\220603 (종평) 용산정수장 현대화 및 고도정수처리시설 설치사업토목, 건축, 기계, 조경분야.xls",
    r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사\2023\입찰결과 - 230206 (종평) 도봉산~옥정 광역철도 3공구 건설공사_Rev.N4.11.xlsb",
    r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사\2024\입찰결과 - 240417 (종심) 시흥거모 공공주택지구 조성공사_Rev.N4.12.xlsb"
]

for sample_file in test_files:
    if not os.path.exists(sample_file):
        continue
    print(f"\nScanning: {sample_file}")
    try:
        if sample_file.endswith('.xlsb'):
            df = pd.read_excel(sample_file, sheet_name=None, engine='pyxlsb', header=None)
        else:
            df = pd.read_excel(sample_file, sheet_name=None, header=None)
            
        for sheet_name, sheet_df in df.items():
            print(f"  Sheet: {sheet_name}")
            found_date = None
            
            # 1. Search for keywords "입찰일시", "개찰일시", "실제 개찰일시"
            for idx, row in sheet_df.head(50).iterrows():
                row_strs = [str(x) for x in row if pd.notna(x)]
                if any("개찰일" in s or "입찰일" in s or "투찰일" in s for s in row_strs):
                    for c_idx, val in enumerate(row):
                        if isinstance(val, str) and ("개찰일" in val or "입찰일" in val or "투찰일" in val):
                            if c_idx + 1 < len(row) and pd.notna(row[c_idx+1]):
                                found_date = row[c_idx+1]
                                print(f"    Found Date Label: '{val}' -> Value: {found_date}")
                                break
                    if found_date: break
                    
            if not found_date:
                # 2. 6자리 파일명 형식 기반 Fallback 확인
                match = re.search(r'(\d{2})(\d{2})(\d{2})', os.path.basename(sample_file))
                if match:
                    print(f"    Date not found in cells. Fallback from filename: {match.group(0)}")
                    
            if found_date:
                break # Move to next file if date logic works out for any sheet
                
    except Exception as e:
        print(f"  Error: {e}")
