import pandas as pd

target_file = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과\입찰결과 - 260113 (종심-고) 새만금항 신항 방파제(연장) 축조공사_Rev.N4.10.xlsb"

def test_extract():
    print("--- 1. Testing 기초정보 sheet R1 ---")
    try:
        df_info = pd.read_excel(target_file, engine='pyxlsb', sheet_name='기초정보', header=None)
        # R1 is row 0, col 17
        r1_val = df_info.iloc[0, 17]
        print(f"R1 value: {r1_val} (Type: {type(r1_val)})")
        
        # Also print what R is exactly in case it's shifted
        print("First row some cols around R (15-20):")
        for i in range(15, min(20, df_info.shape[1])):
            print(f"Col {chr(65+i)}: {df_info.iloc[0, i]}")
            
    except Exception as e:
        print(f"Error reading 기초정보: {e}")

    print("\n--- 2. Testing Main Sheet Col N for Priority ---")
    try:
        # The main sheet is usually the first sheet or named something specific.
        # We will read the first sheet.
        df_main = pd.read_excel(target_file, engine='pyxlsb', header=None)
        
        header_row_idx = -1
        for idx, row in df_main.iterrows():
            if str(row[0]).replace(" ", "") == "순위":
                header_row_idx = idx
                break
                
        if header_row_idx != -1:
            data_df = df_main.iloc[header_row_idx+1:].copy()
            for _, row in data_df.iterrows():
                # column N is index 13
                # Let's print the first few rows to see what's in N
                try:
                    n_val = row[13]
                    company = row[1]
                    print(f"Company: {company}, N_val (index 13): {n_val}")
                except Exception as e:
                    print(f"Row error: {e}")
                    pass
                # Just print top 5
                if _ - (header_row_idx+1) > 5:
                    break
    except Exception as e:
        print(f"Error reading main sheet: {e}")

if __name__ == "__main__":
    test_extract()
