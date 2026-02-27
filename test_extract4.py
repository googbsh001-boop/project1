import pandas as pd

target_file = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과\입찰결과 - 260113 (종심-고) 새만금항 신항 방파제(연장) 축조공사_Rev.N4.10.xlsb"

def test_extract():
    print("--- 1. Testing Main Sheet cells around I6 ---")
    try:
        # I6 means row 5, col 8 (0-indexed)
        # We'll print a grid from row 3 to 10, col 5 to 10
        xls = pd.ExcelFile(target_file, engine='pyxlsb')
        main_sheet_name = [sn for sn in xls.sheet_names if "기초" not in sn][0]
        df_main = pd.read_excel(target_file, engine='pyxlsb', sheet_name=main_sheet_name, header=None)
        
        print("Rows 3 to 10, Cols F(5) to K(10):")
        for r_idx in range(2, 11):
            row_vals = []
            for c_idx in range(5, 12):
                if r_idx < len(df_main) and c_idx < len(df_main.columns):
                    val = df_main.iloc[r_idx, c_idx]
                    row_vals.append(f"{chr(65+c_idx)}{r_idx+1}: {val}")
                else:
                    row_vals.append("N/A")
            print(" | ".join(row_vals))
            
        print("\nAlso checking where '기초금액' and '균형가격' are explicitly mentioned:")
        for r_idx in range(len(df_main)):
            for c_idx in range(len(df_main.columns)):
                val = str(df_main.iloc[r_idx, c_idx]).replace(" ", "")
                if "기초금액" in val or "균형가격" in val or "균형단가" in val:
                    print(f"Found '{val}' at Row {r_idx+1}, Col {chr(65+c_idx)} (Index: {r_idx}, {c_idx})")
                    if c_idx + 1 < len(df_main.columns):
                        print(f" -> Next col value: {df_main.iloc[r_idx, c_idx+1]}")
                    if r_idx + 1 < len(df_main):
                        print(f" -> Next row value: {df_main.iloc[r_idx+1, c_idx]}")
                        
    except Exception as e:
        print(f"Error reading main sheet: {e}")

if __name__ == "__main__":
    test_extract()
