import os
import sys
import pandas as pd

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

def debug_inspect_values():
    test_file = None
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if file.endswith('.xlsb') and not file.startswith('~'):
                test_file = os.path.join(root, file)
                print(f"Using file: {test_file}")
                break
        if test_file:
            break
            
    df = pd.read_excel(test_file, engine='pyxlsb', sheet_name='입찰결과', header=None)
    b_col_data = df.iloc[11:, 1]
    
    valid_bidders = b_col_data.dropna().astype(str).str.strip()
    valid_bidders = valid_bidders[
        (valid_bidders != '') & 
        (valid_bidders.str.lower() != 'nan') & 
        (valid_bidders != '-')
    ]
    
    # Let's inspect the last 50 elements to see what's being counted
    print(f"Length after basic filter: {len(valid_bidders)}")
    print("Last 50 elements:")
    print(valid_bidders.tail(50).to_list())

if __name__ == '__main__':
    debug_inspect_values()
