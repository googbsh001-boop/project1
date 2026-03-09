import os
import pandas as pd

base_path = r"V:\인프라수주팀\인프라자료실\01.입찰결과\입찰결과_09.불참공사"

excel_files = []
for root, dirs, files in os.walk(base_path):
    for file in files:
        if file.endswith(('.xls', '.xlsx')):
            excel_files.append(os.path.join(root, file))

print(f"Total excel files found: {len(excel_files)}")
print("Sample files:")
for f in excel_files[:5]:
    print(f)

if excel_files:
    sample_file = excel_files[0]
    print(f"\nReading sample file: {sample_file}")
    try:
        df = pd.read_excel(sample_file, sheet_name=0)
        print("Columns in first sheet:", df.columns.tolist())
        print("Sample data:")
        print(df.head())
    except Exception as e:
        print(f"Error reading file: {e}")
