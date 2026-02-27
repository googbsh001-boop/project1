import pandas as pd
import os
import re

target_file = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과\입찰결과 - 260113 (종심-고) 새만금항 신항 방파제(연장) 축조공사_Rev.N4.10.xlsb"

def test_extract():
    # Test extracting the note from filename
    filenames = [
        "입찰결과 - 260113 (종심-고) 새만금항 신항 방파제(연장) 축조공사_Rev.N4.10.xlsb",
        "입찰결과 - 241008 (종평) 금호워터폴리스 단지조성공사.xlsb",
        "입찰결과 - 250115 (간이종심) 고양장항 공공주택지구 조경공사 2공구_Rev.N2.17.xlsb"
    ]
    
    for f in filenames:
        match = re.search(r'\((.*?)\)', f)
        if match:
            print(f"Extracted Note for {f}: ({match.group(1)})")
            
    print("\n--- 1. Testing 기초정보 sheet fallback ---")
    try:
        df_info = pd.read_excel(target_file, engine='pyxlsb', sheet_name='기초정보', header=None)
        
        # Look for "입찰마감"
        found = False
        for idx, row in df_info.iterrows():
            for col_idx, cell in enumerate(row):
                if str(cell).replace(" ", "") == "입찰마감":
                    next_val = row[col_idx + 1] if col_idx + 1 < len(row) else None
                    print(f"Found '입찰마감' at row {idx}, col {col_idx}. Next val is: {next_val}")
                    found = True
                    break
            if found: break
            
    except Exception as e:
        print(f"Error reading 기초정보: {e}")

if __name__ == "__main__":
    test_extract()
