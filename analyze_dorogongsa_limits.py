import os
import sys
import pandas as pd
import warnings

# 경고 무시
warnings.filterwarnings('ignore')
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

def analyze_bids(base_dir):
    results = []
    scanned_count = 0
    dorogongsa_count = 0
    
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            file_lower = file.lower()
            if not file.startswith('~') and (file_lower.endswith('.xlsb') or file_lower.endswith('.xlsx') or file_lower.endswith('.xls')):
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(file_path, base_dir)
                scanned_count += 1
                # print(f"검토 중: {rel_path}")
                
                try:
                    if file_lower.endswith('.xlsb'):
                        df = pd.read_excel(file_path, engine='pyxlsb', header=None)
                    else:
                        df = pd.read_excel(file_path, header=None)
                    
                    # 1. 한국도로공사 건인지 확인 (C5 칸 - pandas index로는 row 4, col 2)
                    is_dorogongsa = False
                    try:
                        # C5칸 특화 검사
                        c5_val = str(df.iloc[4, 2]).replace(" ", "")
                        if "한국도로공사" in c5_val:
                            is_dorogongsa = True
                        else:
                            # C5가 아닐 수도 있으니 안전장치로 상단 15x5 범위 한 번 더 체크 (보조)
                            for i in range(min(15, len(df))):
                                for j in range(min(5, len(df.columns))):
                                    cell_val = str(df.iloc[i, j]).replace(" ", "")
                                    if "한국도로공사" in cell_val:
                                        is_dorogongsa = True
                                        break
                                if is_dorogongsa:
                                    break
                    except IndexError:
                        pass
                            
                    if not is_dorogongsa:
                        continue
                        
                    dorogongsa_count += 1
                    print(f"✅ 한국도로공사 입찰 건 발견 ({dorogongsa_count}): [ {rel_path} ]")
                    
                    # 2. 헤더 행 찾기
                    header_row_idx = -1
                    for idx, row in df.iterrows():
                        if str(row[0]).replace(" ", "") == "순위" and str(row[1]).replace(" ", "") == "회사명":
                            header_row_idx = idx
                            break
                            
                    if header_row_idx == -1:
                        print(f"  -> 헤더 오류")
                        continue
                        
                    headers = [str(x).replace("\n", "").replace(" ", "") for x in df.iloc[header_row_idx].values]
                    
                    # 인덱스 찾기
                    def find_col(keywords):
                        for kw in keywords:
                            for c_idx, h in enumerate(headers):
                                if kw in h:
                                    return c_idx
                        return -1
                        
                    idx_company = find_col(['회사명'])
                    idx_yega = find_col(['예가대비'])
                    idx_gicho = find_col(['기초대비'])
                    idx_price_score = find_col(['가격점수'])
                    idx_deduct_score = find_col(['단가감점'])
                    idx_priority = find_col(['우선순위', '낙찰우선순위'])
                    
                    if -1 in [idx_company, idx_yega, idx_price_score, idx_deduct_score]:
                        print(f"  -> 필수 컬럼 누락")
                        continue
                        
                    data_df = df.iloc[header_row_idx+1:]
                    
                    # 3. 50점 만점 데이터 추출
                    valid_bids = []
                    for _, row in data_df.iterrows():
                        try:
                            price_score_raw = str(row[idx_price_score]).strip()
                            deduct_score_raw = str(row[idx_deduct_score]).strip()
                            yega_raw = str(row[idx_yega]).strip()
                            gicho_raw = str(row[idx_gicho]).strip() if idx_gicho != -1 else ""
                            company = str(row[idx_company]).strip()
                            priority_raw = str(row[idx_priority]).strip() if idx_priority != -1 else ""
                            
                            # 데이터 정제
                            if pd.isna(row[idx_price_score]) or pd.isna(row[idx_yega]):
                                continue
                                
                            # 점수 파싱
                            if price_score_raw == 'nan': continue
                            try:
                                p_score = float(price_score_raw)
                            except: continue
                            
                            try:
                                d_score = float(deduct_score_raw) if deduct_score_raw != 'nan' and deduct_score_raw != '-' else 0.0
                            except: d_score = 0.0
                            
                            # 퍼센트 파싱 ("91.898%" -> 91.898)
                            yega_val = float(yega_raw.replace("%", ""))
                            # 1 보다 작으면 100을 곱함 (0.91898 -> 91.898)
                            if yega_val < 2.0:
                                yega_val *= 100
                                
                            valid_bids.append({
                                'company': company,
                                'price_score': p_score,
                                'deduct_score': d_score,
                                'yega_ratio': yega_val,
                                'priority': priority_raw
                            })
                            
                        except Exception as e:
                            pass
                            
                    if not valid_bids:
                        continue
                        
                    # 최고 가격점수(보통 50) 찾기
                    max_price_score = max([b['price_score'] for b in valid_bids])
                    
                    # 감점 없는(max price score & 0 deduction) 최저 예가대비율 찾기
                    perfect_bids = [b for b in valid_bids if b['price_score'] == max_price_score and b['deduct_score'] == 0.0]
                    
                    if perfect_bids:
                        min_yega_bid = min(perfect_bids, key=lambda x: x['yega_ratio'])
                        results.append({
                            'file': file,
                            'winner': min_yega_bid['company'],
                            'max_score': max_price_score,
                            'min_yega_ratio': min_yega_bid['yega_ratio'],
                            'original_priority': min_yega_bid['priority']
                        })
                        print(f"  -> 하한선(예가대비): {min_yega_bid['yega_ratio']:.3f}% ({min_yega_bid['company']})")
                    else:
                        print("  -> 만점/무감점 업체가 없습니다.")
                        
                except Exception as e:
                    print(f"  -> 파일 처리 중 에러: {e}")
                    
    # 결과 요약 통계
    if results:
        print("\n=== 📊 한국도로공사 입찰결과 하한선 통계 요약 ===")
        ratios = [r['min_yega_ratio'] for r in results]
        
        avg_ratio = sum(ratios) / len(ratios)
        min_ratio = min(ratios)
        max_ratio = max(ratios)
        
        print(f"총 {len(results)}건의 도로공사 입찰 분석됨")
        print(f"하한선(예가대비) 최소값 : {min_ratio:.3f}%")
        print(f"하한선(예가대비) 최대값 : {max_ratio:.3f}%")
        print(f"하한선(예가대비) 평균값 : {avg_ratio:.3f}%")
        
        # 분포도
        print("\n[구간별 분포]")
        bins = {
            "89% 미만": 0,
            "89.0 ~ 89.5%": 0,
            "89.5 ~ 90.0%": 0,
            "90.0 ~ 90.5%": 0,
            "90.5 ~ 91.0%": 0,
            "91.0 ~ 91.5%": 0,
            "91.5 ~ 92.0%": 0,
            "92.0% 이상": 0
        }
        
        for r in ratios:
            if r < 89.0: bins["89% 미만"] += 1
            elif r < 89.5: bins["89.0 ~ 89.5%"] += 1
            elif r < 90.0: bins["89.5 ~ 90.0%"] += 1
            elif r < 90.5: bins["90.0 ~ 90.5%"] += 1
            elif r < 91.0: bins["90.5 ~ 91.0%"] += 1
            elif r < 91.5: bins["91.0 ~ 91.5%"] += 1
            elif r < 92.0: bins["91.5 ~ 92.0%"] += 1
            else: bins["92.0% 이상"] += 1
            
        for k, v in bins.items():
            if v > 0:
                print(f" - {k}: {v}건")
                
if __name__ == "__main__":
    analyze_bids(BASE_DIR)
