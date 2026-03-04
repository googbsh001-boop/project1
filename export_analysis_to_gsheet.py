import gspread
from google.oauth2.service_account import Credentials
import sys

# 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1n3WxFMxjS-mhHGE8I4dXi4Q2oJ3l4sq_OkCkBeJkbJI'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_google_sheet_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def main():
    client = get_google_sheet_client()
    sh = client.open_by_key(SHEET_ID)
    
    worksheet_name = "우선순위 분석 결과"
    
    try:
        worksheet = sh.worksheet(worksheet_name)
        worksheet.clear()
        print(f"기존 '{worksheet_name}' 시트를 초기화했습니다.")
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sh.add_worksheet(title=worksheet_name, rows="100", cols="20")
        print(f"새로운 '{worksheet_name}' 시트를 생성했습니다.")
    
    data = [
        ["공사명(공구)", "우선순위 1위 업체", "입찰금액(원)", "예가대비(%)", "기초대비(%)", "가격점수", "단가감점상태", "1위 선정 사유 (분석결과)"],
        ["계양-강화 제4공구", "금호건설 주식회사", "167,646,686,770", "91.898%", "92.241%", "50.000", "0.000", "가격점수 만점(50점) 및 단가감점 없음(0점)을 충족하는 범위 내에서 최저가 투찰"],
        ["계양-강화 제3공구", "동부건설 주식회사", "154,764,864,282", "91.099%", "91.145%", "50.000", "0.000", "가격점수 만점(50점) 및 단가감점 없음(0점)을 충족하는 범위 내에서 최저가 투찰"],
        ["계양-강화 제5공구", "현대건설 주식회사", "212,104,661,887", "89.640%", "88.792%", "50.000", "0.000", "가격점수 만점(50점) 및 단가감점 없음(0점)을 충족하는 범위 내에서 최저가 투찰"],
        [],
        ["<종합 분석 요약>"],
        ["1. 방식 설명", "해당 공사들은 '종합심사낙찰제(일반공사)' 방식으로 진행되었습니다."],
        ["2. 1위의 조건", "단순히 순위(입찰금액이 낮은 순)대로 낙찰되는 것이 아니라, '가격점수 만점(50점)'과 '단가감점(0점)' 기준을 동시에 통과해야 합니다."],
        ["3. 최종 결정", "이 두 가지 심사 요건을 완벽하게 통과한 업체들 중에서 가장 낮은 입찰금액(초저가 방어선을 지킨 최저가)을 적어낸 업체가 최종적으로 우선순위 1위를 차지하게 되었습니다."],
        ["", "즉, 가격점수를 깎이지 않는 투찰하한선을 정확히 예측하고 그 선에 가장 가깝게 쓴 회사가 1위가 된 것입니다."],
        [],
        ["<참여 업체 투찰 전략 분석>"],
        ["전략 그룹", "특징 및 공통점", "결과"],
        ["1. 정밀 타격형 (낙찰 그룹)", "가격점수 만점(50점)과 단가감점 0점을 받을 수 있는 '최저 하한선(투찰마지노선)'을 정확히 예측하여 투찰. (예: 4공구 금호 91.898%, 3공구 동부 91.099%, 5공구 현대 89.640%)", "모든 심사 요건을 완벽히 충족하며 최저가로 1위 선점."],
        ["2. 공격적 저가형 (1순위 밖 상위권)", "낙찰 그룹보다 더 낮은 금액을 써내어 단순 가격 순위로는 앞서지만, 이로 인해 가격점수에서 미세한 감점(예: 49.999점, 49.998점 등)을 받음.", "0.001점의 미세한 감점이라도, 만점자가 존재하는 한 종합심사에서 치명적으로 작용하여 우선순위에서 밀려남."],
        ["3. 보수적 안전형 (중하위권)", "단가감점이나 가격점수 감점을 피하기 위해 하한선보다 넉넉히 높은 금액으로 투찰.", "점수는 만점(50점)을 받지만, 가격 경쟁력에서 밀려 우선순위를 확보하지 못함."],
        ["결론", "종합심사낙찰제에서는 단순 저가 입찰이 아니라, '감점 없는 최저가'를 찾는 고도의 눈치싸움과 통계적 예측이 핵심 승부처임."]
    ]
    
    worksheet.update(values=data, range_name='A1')
    
    # 간단한 서식 적용 (헤더 볼드)
    try:
        from gspread_formatting import CellFormat, TextFormat, format_cell_range
        fmt_header = CellFormat(textFormat=TextFormat(bold=True))
        format_cell_range(worksheet, 'A1:H1', fmt_header)
        format_cell_range(worksheet, 'A6', fmt_header)
        format_cell_range(worksheet, 'A12', fmt_header)
        format_cell_range(worksheet, 'A13:C13', fmt_header)
    except Exception as e:
        print(f"서식 적용 생략: {e}")
        
    print("구글 시트에 분석 결과 업데이트 완료!")

if __name__ == "__main__":
    main()
