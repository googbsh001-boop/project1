import os
import sys
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
base_path = r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사"

excel_files = []
for root, dirs, files in os.walk(base_path):
    for file in files:
        if file.endswith(('.xls', '.xlsx', '.xlsb')) and not file.startswith('~'):
            excel_files.append(os.path.join(root, file))

# Limit to first 20 for debugging
for fpath in excel_files[:20]:
    print(f"\nScanning: {fpath}")
    try:
        if fpath.endswith('.xlsb'):
            df = pd.read_excel(fpath, sheet_name=0, engine='pyxlsb', header=None)
        else:
            df = pd.read_excel(fpath, sheet_name=0, header=None)
            
        found = False
        for idx, row in df.iterrows():
            row_strs = [str(x) for x in row if pd.notna(x)]
            if any("진흥" in s for s in row_strs):
                print(f"  Found '진흥' in row {idx}:")
                for c_idx, val in enumerate(row):
                    if pd.notna(val):
                        # print non-null columns
                        col_letter = chr(65 + c_idx) if c_idx < 26 else chr(64 + c_idx // 26) + chr(65 + c_idx % 26)
                        print(f"    Col {col_letter} (idx {c_idx}): {val}")
                found = True
        if not found:
            print("  '진흥' not found in this file.")
    except Exception as e:
        print(f"  Error reading {fpath}: {e}")
