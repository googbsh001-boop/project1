import os
import pandas as pd

def extract_hanwha_data(base_dir):
    result = []
    
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.endswith('.xlsb') and not file.startswith('~'):
                file_path = os.path.join(root, file)
                print(f"Processing: {file_path}")
                try:
                    df = pd.read_excel(file_path, engine='pyxlsb', header=None)
                    
                    # 1. 시트에서 "순위"나 "회사명"이 있는 헤더 행 찾기
                    header_row_idx = -1
                    for idx, row in df.iterrows():
                        if str(row[0]).replace(" ", "") == "순위" and str(row[1]).replace(" ", "") == "회사명":
                            header_row_idx = idx
                            break
                    
                    if header_row_idx == -1:
                        print(f"  -> Could not find header row in {file}")
                        continue
                        
                    # Find indices for target columns from the header row
                    header_row = [str(x).replace("\n", "").replace(" ", "") for x in df.iloc[header_row_idx].values]
                    
                    try:
                        idx_company = header_row.index('회사명')
                        # '기초대비' or '투찰율'
                        idx_ratio = -1
                        if '기초대비' in header_row:
                            idx_ratio = header_row.index('기초대비')
                        elif '투찰율' in header_row:
                            idx_ratio = header_row.index('투찰율')
                            
                        # '낙찰우선순위' or '우선순위' or fallback to '순위'
                        idx_priority = -1
                        if '낙찰우선순위' in header_row:
                            idx_priority = header_row.index('낙찰우선순위')
                        elif '우선순위' in header_row:
                            idx_priority = header_row.index('우선순위')
                        elif '순위' in header_row:
                            idx_priority = header_row.index('순위')
                            
                    except ValueError as e:
                        print(f"  -> Missing required columns in header: {e}")
                        continue

                    # Search for Hanwha in data rows
                    data_df = df.iloc[header_row_idx+1:]
                    participated = "X"
                    ratio_val = ""
                    priority_val = ""
                    
                    for _, row in data_df.iterrows():
                        company = str(row[idx_company]).strip()
                        if "한화" in company:
                            participated = "O"
                            ratio_val = row[idx_ratio] if idx_ratio != -1 and not pd.isna(row[idx_ratio]) else ""
                            priority_val = row[idx_priority] if idx_priority != -1 and not pd.isna(row[idx_priority]) else ""
                            break
                            
                    result.append({
                        "파일명": file,
                        "입찰유무": participated,
                        "기초대비투찰율": ratio_val,
                        "낙찰우선순위": priority_val
                    })
                    print(f"  -> {participated}, ratio: {ratio_val}, priority: {priority_val}")
                except Exception as e:
                    print(f"  -> Error: {e}")
                    
    return result

if __name__ == "__main__":
    base_dir = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"
    data = extract_hanwha_data(base_dir)
    print("\n--- Summary ---")
    for d in data:
        print(d)
