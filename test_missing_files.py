import os
import glob
import pandas as pd

folder = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"
all_files = glob.glob(os.path.join(folder, "**", "*.xlsb"), recursive=True) + glob.glob(os.path.join(folder, "**", "*.xlsx"), recursive=True)
files = [f for f in all_files if not os.path.basename(f).startswith('~$')]

processed = pd.read_csv("bidding_analysis.csv")['File'].unique()
missing = [f for f in files if os.path.basename(f) not in processed]

for f in missing[:3]:
    print(f"File: {os.path.basename(f)}")
    try:
        df = pd.read_excel(f, sheet_name="입찰결과", header=None, engine='pyxlsb' if f.endswith('xlsb') else 'openpyxl')
        for i, row in df.head(40).iterrows():
            row_str = "".join([str(x) for x in row.values if pd.notna(x)])
            if any(keyword in row_str for keyword in ['업체', '대표사', '회사', '투찰', '사정률', '기초']):
                print(f"Row {i} match: {row.values}")
    except Exception as e:
        print(f"Error loading: {e}")
    print("-" * 50)
