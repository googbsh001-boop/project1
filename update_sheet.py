import gspread
from google.oauth2.service_account import Credentials
import datetime

# 사용할 스코프 설정
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def update_google_sheet(credentials_file, sheet_id, worksheet_name, data):
    """
    구글 스프레드시트에 데이터를 추가하는 함수
    """
    try:
        print(f"인증 시도 중... 파일: {credentials_file}")
        # 인증 처리
        credentials = Credentials.from_service_account_file(
            credentials_file, scopes=SCOPES
        )
        gc = gspread.authorize(credentials)
        print("인증 성공")

        # 스프레드시트 열기
        print(f"스프레드시트 여는 중... ID: {sheet_id}")
        sh = gc.open_by_key(sheet_id)

        # 워크시트 선택 (없으면 첫 번째 시트 선택)
        try:
            worksheet = sh.worksheet(worksheet_name)
            print(f"워크시트 '{worksheet_name}' 선택됨")
        except gspread.WorksheetNotFound:
            print(f"워크시트 '{worksheet_name}'를 찾을 수 없어 첫 번째 시트를 사용합니다.")
            worksheet = sh.sheet1

        # 데이터 쓰기 (append_row는 다음 빈 행에 추가)
        # print(f"데이터 추가 중: {data}")
        # worksheet.append_row(data)
        # print("데이터가 성공적으로 추가되었습니다.")
        
        # 특정 셀 업데이트
        cell_address = 'A5'
        value = 10
        print(f"셀 {cell_address}에 값 {value} 쓰는 중...")
        worksheet.update_acell(cell_address, value)
        print("셀 업데이트 성공!")

    except FileNotFoundError:
        print(f"오류: '{credentials_file}' 파일을 찾을 수 없습니다. 경로를 확인해주세요.")
    except Exception as e:
        print(f"오류가 발생했습니다: {e}")

if __name__ == "__main__":
    # --- 설정 필요 ---
    # 1. 구글 클라우드 콘솔에서 다운로드 받은 서비스 계정 키 파일의 이름이 'credentials.json'이어야 합니다.
    CREDENTIALS_FILE = 'credentials.json' 
    
    # 2. 공유하려는 구글 스프레드시트의 URL에서 ID 부분을 복사해서 아래에 넣어주세요.
    # 예: https://docs.google.com/spreadsheets/d/abc12345/edit 중 'abc12345' 부분
    SHEET_ID = '1ZwmfkDFJDBLK_U2oymeI3XZ3w5FLLak7AxodXN1pXrU' 
    
    # 3. 데이터를 입력할 시트 이름 (기본값: Sheet1)
    WORKSHEET_NAME = 'Sheet1' 

    # --- 테스트 데이터 ---
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sample_data = [current_time, "테스트 데이터", "자동 입력 성공"]
    
    if SHEET_ID == '여기에_스프레드시트_ID를_입력하세요':
        print("경고: 스크립트 내의 SHEET_ID를 실제 스프레드시트 ID로 변경해주세요.")
    else:
        update_google_sheet(CREDENTIALS_FILE, SHEET_ID, WORKSHEET_NAME, sample_data)
