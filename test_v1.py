import os
import pandas as pd

BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

for root, dirs, files in os.walk(BASE_DIR):
    for file in files:
        if file.startswith("입찰결과") and file.endswith(".xlsb"):
            sample_file = os.path.join(root, file)
            print(f"Testing {sample_file}")
            
            xls = pd.ExcelFile(sample_file, engine='pyxlsb')
            main_sheet = [sn for sn in xls.sheet_names if "기초" not in sn][0]
            
            df = pd.read_excel(sample_file, engine='pyxlsb', sheet_name=main_sheet, header=None)
            
            # V1 is row 0, col 21
            v1_main = df.iloc[0, 21] if df.shape[1] > 21 else "No col V in main"
            print(f"V1 in main sheet: {v1_main}")
            
            try:
                df_info = pd.read_excel(sample_file, engine='pyxlsb', sheet_name='기초정보', header=None)
                v1_info = df_info.iloc[0, 21] if df_info.shape[1] > 21 else "No col V in info"
                print(f"V1 in 기초정보: {v1_info}")
            except Exception as e:
                print(f"Error info sheet: {e}")
                
            break
    break
