import os
import pandas as pd
import sys

# BASE_DIR = r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사"
# For faster testing, only look at recent years
BASE_DIR = r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사"

def check_blanks():
    excel_files = []
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if file.endswith(('.xls', '.xlsx', '.xlsb')) and not file.startswith('~'):
                if "2020전" in root or "2021" in root or "2022" in root or "2023" in root:
                    continue
                excel_files.append((root, file))
                
    print(f"Total files: {len(excel_files)}")
    
    for root, file in excel_files:
        fpath = os.path.join(root, file)
        try:
            if file.endswith('.xlsb'):
                df_all = pd.read_excel(fpath, sheet_name=None, engine='pyxlsb', header=None)
                df = df_all[list(df_all.keys())[0]]
            else:
                df_all = pd.read_excel(fpath, sheet_name=None, header=None)
                df = df_all[list(df_all.keys())[0]]
                
            for idx, row in df.iterrows():
                row_strs = [str(x) for x in row if pd.notna(x)]
                if any("진흥" in s for s in row_strs):
                    n_val = row[13] if len(row) > 13 else None
                    rank_val = "" if pd.isna(n_val) else str(n_val).strip()
                    
                    bid_amount = row[2] if len(row) > 2 else ""
                    bid_amount = "" if pd.isna(bid_amount) else str(bid_amount).strip()
                    
                    if rank_val == "" or rank_val == "nan" or bid_amount == "" or bid_amount == "nan":
                        print(f"\n--- MISSING DATA IN : {file} ---")
                        print(f"Row {idx}: {row.tolist()}")
                        
        except Exception as e:
            pass

if __name__ == "__main__":
    check_blanks()
