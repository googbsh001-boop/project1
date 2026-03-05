import os
import pandas as pd
import glob

folder = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"
all_files = glob.glob(os.path.join(folder, "**", "*.xlsb"), recursive=True) + glob.glob(os.path.join(folder, "**", "*.xlsx"), recursive=True)
files = [f for f in all_files if not os.path.basename(f).startswith('~$')]

if files:
    test_file = files[0]
    print(f"Reading: {test_file}")
    df = pd.read_excel(test_file, sheet_name="입찰결과", header=None, engine='pyxlsb')
    print("Searching for header row...")
    for i, row in df.head(20).iterrows():
        row_list = [str(x).strip().replace('\n', '').replace(' ', '') for x in row.values]
        if any('업체' in x or '대표사' in x for x in row_list) or any('투찰' in x for x in row_list):
            print(f"Row {i}: {row.values}")
else:
    print("No valid files found.")
