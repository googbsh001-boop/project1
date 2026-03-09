import os
import pandas as pd
import sys
import re

BASE_DIR = r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사"

def check_valid_bids():
    excel_files = []
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if file.endswith(('.xls', '.xlsx', '.xlsb')) and not file.startswith('~'):
                if "2020전" in root or "2021" in root or "2022" in root or "2023" in root:
                    continue
                excel_files.append((root, file))
                
    actual_bids = []
    garbage_bids = []
    
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
                    bid_ratio = row[4] if len(row) > 4 else ""
                    
                    try: rank_val = str(int(float(rank_val)))
                    except: pass
                    
                    if pd.isna(bid_amount) or str(bid_amount).strip() == "":
                        # This is likely garbage or shifted
                        garbage_bids.append((file, idx, row.tolist()))
                    else:
                        actual_bids.append((file, idx, rank_val, bid_amount, bid_ratio, row.tolist()))
                        
        except Exception as e:
            pass

    print(f"Total valid bids: {len(actual_bids)}")
    print(f"Total garbage bids: {len(garbage_bids)}")
    
    print("\n--- SAMPLE GARBAGE BIDS (Empty Bid Amount) ---")
    for b in garbage_bids[:10]:
        print(f"{b[0]} Row {b[1]}: {b[2][:15]}")

    print("\n--- SAMPLE VALID BIDS (Has Bid amount) ---")
    count_missing_rank = 0
    count_missing_ratio = 0
    for b in actual_bids:
        if b[2] == "" or b[2].lower() == "nan":
            count_missing_rank += 1
            print(f"Missing Rank in {b[0]} Row {b[1]}: rank='{b[2]}', amt='{b[3]}', ratio='{b[4]}', row={b[5][:15]}")
        if str(b[4]).strip() == "" or str(b[4]).lower() == "nan":
            count_missing_ratio += 1
            # print(f"Missing Ratio in {b[0]} Row {b[1]}: rank='{b[2]}', amt='{b[3]}', ratio='{b[4]}'")

    print(f"\nMissing Rank valid bids: {count_missing_rank}")
    print(f"Missing Ratio valid bids: {count_missing_ratio}")
    

if __name__ == "__main__":
    check_valid_bids()
