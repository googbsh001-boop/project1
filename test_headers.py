import os
import pandas as pd
import glob

folder = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"
all_files = glob.glob(os.path.join(folder, "**", "*.xlsb"), recursive=True) + glob.glob(os.path.join(folder, "**", "*.xlsx"), recursive=True)
files = [f for f in all_files if not os.path.basename(f).startswith('~$')]

for f in files[:2]:
    print(f"File: {os.path.basename(f)}")
    df = pd.read_excel(f, sheet_name="입찰결과", header=None, engine='pyxlsb' if f.endswith('xlsb') else 'openpyxl')
    header_idx = -1
    for i, row in df.head(30).iterrows():
        row_str = "".join([str(x) for x in row.values if pd.notna(x)])
        if '업체' in row_str or '대표사' in row_str:
            header_idx = i
            break
            
    if header_idx != -1:
        headers = df.iloc[header_idx].values
        clean_headers = [str(x).replace('\n', '') for x in headers]
        print(f"Headers found at row {header_idx}:")
        print(clean_headers)
    else:
        print("Header not found.")
    print("-" * 50)
