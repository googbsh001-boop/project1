import os
import sys
import pandas as pd

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

def debug_one():
    test_file = None
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if file.endswith('.xlsb') and not file.startswith('~'):
                test_file = os.path.join(root, file)
                print(f"Using file: {file}")
                break
        if test_file:
            break
            
    if not test_file:
        print("No files found.")
        return

    df = pd.read_excel(test_file, engine='pyxlsb', sheet_name='입찰결과', header=None)
    
    b_col_data = df.iloc[11:, 1]
    valid_bidders = b_col_data.dropna().astype(str).str.strip()
    valid_bidders = valid_bidders[(valid_bidders != '') & (valid_bidders.str.lower() != 'nan')]
    
    print(f"Total rows in df: {len(df)}")
    print(f"Number of valid bidders: {len(valid_bidders)}")
    print("Top 5 bidders:")
    print(valid_bidders.head())
    print("\nBottom 10 bidders:")
    print(valid_bidders.tail(10))

if __name__ == '__main__':
    debug_one()
