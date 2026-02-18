import sys
import os
import pandas as pd

# 파일 경로 설정
folder_path = r"E:\인프라수주팀\입찰결과분석"
output_file = "excel_preview.txt"

try:
    files = os.listdir(folder_path)
    # xlsb 파일 찾기
    target_files = [f for f in files if f.endswith('.xlsb')]
    
    if not target_files:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write("No .xlsb files found")
        sys.exit(1)
        
    target_file = target_files[0]
    file_path = os.path.join(folder_path, target_file)

    # 엑셀 파일 읽기 (헤더 없이 읽어서 구조 확인)
    df = pd.read_excel(file_path, engine='pyxlsb', header=None)
    
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(f"File: {target_file}\n")
        f.write("-" * 50 + "\n")
        f.write("First 20 rows:\n")
        f.write(df.head(20).to_string())
        f.write("\n" + "-" * 50 + "\n")

    print(f"Preview saved to {output_file}")

except Exception as e:
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(f"Error: {e}")
    print(f"Error occurred: {e}")
