import pandas as pd
import os

# 파일 경로 설정
folder_path = r"E:\인프라수주팀\입찰결과분석"
files = os.listdir(folder_path)
target_file = [f for f in files if f.endswith('.xlsb')][0] # 첫 번째 파일 선택
file_path = os.path.join(folder_path, target_file)

print(f"분석 대상 파일: {file_path}")

try:
    # 엑셀 파일 읽기 (엔진: pyxlsb)
    df = pd.read_excel(file_path, engine='pyxlsb')
    
    print("\n--- 데이터프레임 정보 ---")
    print(df.info())
    
    print("\n--- 상위 5개 행 출력 ---")
    print(df.head())
    
    print("\n--- 컬럼 목록 ---")
    print(df.columns.tolist())

except Exception as e:
    print(f"오류 발생: {e}")
