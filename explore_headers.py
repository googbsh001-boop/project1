import os
import sys
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
sample_file = r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사\2020전\170103 (종심-일)대청댐계통 광역상수도사업 제2공구 정수시설공사.xlsx"

try:
    df = pd.read_excel(sample_file, sheet_name='입력(개찰후)', header=None)
    for idx in range(25, 35):
        row = df.iloc[idx].tolist()
        # Print column indices with their values
        print(f"\nRow {idx}:")
        for i, val in enumerate(row):
            if pd.notna(val):
                # Excel column letter (A, B, C...)
                col_letter = chr(65 + i) if i < 26 else chr(64 + i // 26) + chr(65 + i % 26)
                print(f"  {col_letter} (idx {i}): {val}")
except Exception as e:
    print(e)
