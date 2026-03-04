import os
import sys

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

def debug_keys():
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if file.endswith('.xlsb') and not file.startswith('~'):
                key = file.replace('.xlsb', '').strip()
                print(f"Key: {key}")
                
if __name__ == '__main__':
    debug_keys()
