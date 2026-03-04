import sys
import re
import gspread
from google.oauth2.service_account import Credentials

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1n3WxFMxjS-mhHGE8I4dXi4Q2oJ3l4sq_OkCkBeJkbJI'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def extract_date(filename):
    """
    파일명에서 6자리 입찰날짜(YYMMDD)를 추출합니다.
    예: '입찰결과 - 260113 (종심-고) 새만금항 신항 방파제(연장) 축조공사_Rev.N4.10.xlsb' -> '260113'
    """
    # 숫자 6자리가 연속으로 나오는 부분을 찾습니다.
    match = re.search(r'(\d{6})', filename)
    if match:
        return match.group(1)
    return "999999"  # 날짜가 없으면 맨 뒤로 보내기 위해 큰 값 설정

def sort_sheet():
    print("구글 시트 인증 시도...")
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)
    
    try:
        # worksheet ID: 1426009222 -> '시트2'
        worksheet = sh.get_worksheet_by_id(1426009222)
        print(f"'{worksheet.title}' 시트를 찾았습니다. 데이터를 읽어옵니다.")
        
        # 전체 데이터 읽기
        all_data = worksheet.get_all_values()
        if not all_data or len(all_data) <= 1:
            print("데이터가 부족하여 정렬을 수행하지 않습니다.")
            return

        headers = all_data[0]
        rows = all_data[1:]
        
        # A열(인덱스 0, '공고명')의 문자열에서 날짜를 추출하여 정렬 기준으로 사용
        # 오름차순 정렬
        sorted_rows = sorted(rows, key=lambda x: extract_date(x[0]))
        
        # 정렬된 데이터로 시트 업데이트
        final_data = [headers] + sorted_rows
        
        print(f"{len(sorted_rows)}개의 행을 입찰날짜 오름차순으로 새로 업데이트합니다...")
        worksheet.clear()
        worksheet.update('A1', final_data)
        
        print("정렬 완료!")
        
    except Exception as e:
        print(f"에러 발생: {e}")

if __name__ == "__main__":
    sort_sheet()
