import os
import sys
import pandas as pd

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

def debug_counts():
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if file.endswith('.xlsb') and not file.startswith('~'):
                file_path = os.path.join(root, file)
                
                try:
                    try:
                        df = pd.read_excel(file_path, engine='pyxlsb', sheet_name='입찰결과', header=None)
                    except Exception as e:
                        df = pd.read_excel(file_path, engine='pyxlsb', header=None)
                    
                    if len(df.columns) > 1 and len(df) > 11:
                        b_col_data = df.iloc[11:, 1]
                        valid_bidders = b_col_data.dropna().astype(str).str.strip()
                        # Filter out empty strings and "nan" strings that pandas might produce
                        valid_bidders = valid_bidders[(valid_bidders != '') & (valid_bidders.str.lower() != 'nan')]
                        count = len(valid_bidders)
                        print(f"File: {file} -> Count: {count}")
                    else:
                        print(f"File: {file} -> Count: 0 (Not enough rows/cols)")
                        
                except Exception as e:
                    print(f"File: {file} -> Error: {e}")

if __name__ == '__main__':
    debug_counts()
