import os
import sys
import pandas as pd
import warnings
from collections import defaultdict
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import CellFormat, TextFormat, format_cell_range

# 경고 무시
warnings.filterwarnings('ignore')
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"
OUTPUT_EXCEL = r"E:\인프라수주팀\트레이닝\프로젝트1\타겟팅_시뮬레이션_결과.xlsx"

# 구글 시트 연동 정보
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

def simulate_and_export(base_dir):
    bidding_data = [] 
    target_stats = defaultdict(lambda: {
        'total_encounters': 0, 
        'success_wins': 0,
        'fail_deduction': 0,
        'fail_too_high': 0
    })
    
    scanned_count = 0
    dorogongsa_count = 0
    
    print("도로공사 데이터 파싱 시작...")
    
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            file_lower = file.lower()
            if not file.startswith('~') and (file_lower.endswith('.xlsb') or file_lower.endswith('.xlsx') or file_lower.endswith('.xls')):
                file_path = os.path.join(root, file)
                
                try:
                    if file_lower.endswith('.xlsb'):
                        df = pd.read_excel(file_path, engine='pyxlsb', header=None)
                    else:
                        df = pd.read_excel(file_path, header=None)
                    
                    is_dorogongsa = False
                    try:
                        c5_val = str(df.iloc[4, 2]).replace(" ", "")
                        if "한국도로공사" in c5_val:
                            is_dorogongsa = True
                        else:
                            for i in range(min(15, len(df))):
                                for j in range(min(5, len(df.columns))):
                                    cell_val = str(df.iloc[i, j]).replace(" ", "")
                                    if "한국도로공사" in cell_val:
                                        is_dorogongsa = True
                                        break
                                if is_dorogongsa: break
                    except IndexError: pass
                            
                    if not is_dorogongsa: continue
                    dorogongsa_count += 1
                    
                    header_row_idx = -1
                    for idx, row in df.iterrows():
                        if str(row[0]).replace(" ", "") == "순위" and str(row[1]).replace(" ", "") == "회사명":
                            header_row_idx = idx
                            break
                    if header_row_idx == -1: continue
                        
                    headers = [str(x).replace("\n", "").replace(" ", "") for x in df.iloc[header_row_idx].values]
                    def find_col(kws):
                        for kw in kws:
                            for c_idx, h in enumerate(headers):
                                if kw in h: return c_idx
                        return -1
                        
                    idx_company = find_col(['회사명'])
                    idx_yega = find_col(['예가대비'])
                    idx_price_score = find_col(['가격점수'])
                    idx_deduct_score = find_col(['단가감점'])
                    
                    if -1 in [idx_company, idx_yega, idx_price_score, idx_deduct_score]: continue
                        
                    data_df = df.iloc[header_row_idx+1:]
                    
                    valid_bids = []
                    for _, row in data_df.iterrows():
                        try:
                            price_score_raw = str(row[idx_price_score]).strip()
                            deduct_score_raw = str(row[idx_deduct_score]).strip()
                            yega_raw = str(row[idx_yega]).strip()
                            company = str(row[idx_company]).strip()
                            
                            if pd.isna(row[idx_price_score]) or pd.isna(row[idx_yega]) or price_score_raw == 'nan':
                                continue
                                
                            try: p_score = float(price_score_raw)
                            except: continue
                            
                            try: d_score = float(deduct_score_raw) if deduct_score_raw not in ['nan', '-'] else 0.0
                            except: d_score = 0.0
                            
                            yega_val = float(yega_raw.replace("%", ""))
                            if yega_val < 2.0: yega_val *= 100
                                
                            valid_bids.append({
                                'company': company,
                                'price_score': p_score,
                                'deduct_score': d_score,
                                'yega_ratio': yega_val
                            })
                        except: pass
                            
                    if not valid_bids: continue
                        
                    max_price_score = max([b['price_score'] for b in valid_bids])
                    perfect_bids = [b for b in valid_bids if b['price_score'] == max_price_score and b['deduct_score'] == 0.0]
                    
                    if perfect_bids:
                        min_yega_bid = min(perfect_bids, key=lambda x: x['yega_ratio'])
                        limit_yega = min_yega_bid['yega_ratio']
                        
                        bidding_data.append({
                            'limit_yega': limit_yega,
                            'bids': valid_bids
                        })

                except Exception as e:
                    print(f"에러: {e}")
                    
    print(f"데이터 파싱 완료 (도로공사 건 총 {dorogongsa_count}개)\n시뮬레이터 가동 중...")
    
    # 전략 시뮬레이션
    target_margin = 0.001 
    
    for data in bidding_data:
        limit_yega = data['limit_yega']
        bids = data['bids']
        
        # 완벽 투찰(1위/만점/무감점) 리스트 생성
        max_price_score = max([b['price_score'] for b in bids])
        perfect_bids = [b for b in bids if b['price_score'] == max_price_score and b['deduct_score'] == 0.0]
        
        for target_bid in bids:
            target_comp = target_bid['company']
            my_sim_yega = target_bid['yega_ratio'] - target_margin
            
            target_stats[target_comp]['total_encounters'] += 1
            
            if my_sim_yega < limit_yega:
                target_stats[target_comp]['fail_deduction'] += 1
            else:
                is_first = True
                for other_bid in perfect_bids:
                    if limit_yega <= other_bid['yega_ratio'] < my_sim_yega:
                        is_first = False
                        break
                
                if is_first:
                    target_stats[target_comp]['success_wins'] += 1
                else:
                    target_stats[target_comp]['fail_too_high'] += 1

    # 결과 표 생성을 위한 데이터 가공 (3회 이상)
    valid_targets = {k: v for k, v in target_stats.items() if v['total_encounters'] >= 3}
    sorted_targets = sorted(
        valid_targets.items(), 
        key=lambda x: (x[1]['success_wins'] / x[1]['total_encounters']), 
        reverse=True
    )
    
    # 엑셀과 구글시트에 쓸 데이터 배열 만들기
    gsheet_data = [
        ["타겟팅 시뮬레이션 결과 (기준: -0.001%)"],
        ["순위", "타겟 건설사", "시뮬레이션 승률", "조우 횟수", "1순위 낙찰", "감점 탈락", "보수적 탈락", "전략적 분석(코멘트)"]
    ]
    
    excel_data = []

    for i, (comp, stats) in enumerate(sorted_targets):
        total = stats['total_encounters']
        wins = stats['success_wins']
        fail_deduct = stats['fail_deduction']
        fail_high = stats['fail_too_high']
        
        win_rate = (wins / total) * 100
        win_rate_str = f"{win_rate:.1f}%"
        deduct_rate = (fail_deduct / total) * 100
        
        if win_rate > 30:
            msg = f"훌륭한 타겟: 0.001% 낮게 써도 감점 확률 {deduct_rate:.1f}% 불과, 1위 확률 매우 높음"
        elif deduct_rate > 50:
            msg = f"위험 타겟(회피 권장): 타겟팅 시 {deduct_rate:.1f}% 확률로 동반 감점 자폭"
        else:
            msg = f"무난/보수적 타겟: 내가 더 낮게 쓰더라도 제3자에게 질 확률이 높음"
            
        gsheet_data.append([
            i+1, comp, win_rate_str, total, wins, fail_deduct, fail_high, msg
        ])
        
        excel_data.append({
            "순위": i+1,
            "타겟 건설사": comp,
            "시뮬레이션 승률": win_rate_str,
            "조우 횟수": total,
            "1순위 낙찰": wins,
            "감점 탈락": fail_deduct,
            "보수적 탈락": fail_high,
            "전략적 분석": msg
        })

    # ==========================
    # 1. 엑셀 파일로 내보내기
    # ==========================
    print("\n엑셀 파일 추출 중...")
    try:
        df_export = pd.DataFrame(excel_data)
        df_export.to_excel(OUTPUT_EXCEL, index=False)
        print(f"✅ 엑셀 저장 완료: {OUTPUT_EXCEL}")
    except Exception as e:
        print(f"❌ 엑셀 저장 실패: {e}")

    # ==========================
    # 2. 구글 시트로 업데이트
    # ==========================
    print("\n구글 시트 연동 중...")
    try:
        client = get_google_sheet_client()
        sh = client.open_by_key(SHEET_ID)
        
        worksheet_name = "타겟팅 시뮬레이션 결과"
        
        try:
            worksheet = sh.worksheet(worksheet_name)
            worksheet.clear()
            print(f"기존 '{worksheet_name}' 시트를 초기화했습니다.")
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=worksheet_name, rows="100", cols="10")
            print(f"새로운 '{worksheet_name}' 시트를 생성했습니다.")
        
        # 데이터 삽입
        worksheet.update(values=gsheet_data, range_name='A1')
        
        # 간단 서식 (헤더 강조)
        try:
            fmt_title = CellFormat(textFormat=TextFormat(bold=True, fontSize=12))
            fmt_header = CellFormat(backgroundColor={'red': 0.9, 'green': 0.9, 'blue': 0.9}, 
                                    textFormat=TextFormat(bold=True))
            format_cell_range(worksheet, 'A1:H1', fmt_title)
            format_cell_range(worksheet, 'A2:H2', fmt_header)
        except Exception as e:
            print(f"서식 적용 생략: {e}")
            
        print("✅ 구글 시트 업데이트 완료!")
        
    except Exception as e:
        print(f"❌ 구글 시트 연동 실패: {e}")

if __name__ == "__main__":
    simulate_and_export(BASE_DIR)
