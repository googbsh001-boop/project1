import os
import sys
import pandas as pd
import warnings
from collections import defaultdict

# 경고 무시
warnings.filterwarnings('ignore')
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

def analyze_bids(base_dir):
    results = []
    company_stats = defaultdict(lambda: {
        'total_bids': 0, 
        'wins': 0, # 만점 하한선 1위 횟수
        'avg_diff_from_limit': 0.0, # 하한선과의 평균 차이 (%)
        'diffs': [],
        'aggresive_count': 0, # 공격적 저가 (하한선 미만 타격으로 감점) 횟수
        'conservative_count': 0 # 보수적 투찰 (하한선 대비 너무 높음) 횟수
    })
    
    scanned_count = 0
    dorogongsa_count = 0
    
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            file_lower = file.lower()
            if not file.startswith('~') and (file_lower.endswith('.xlsb') or file_lower.endswith('.xlsx') or file_lower.endswith('.xls')):
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(file_path, base_dir)
                scanned_count += 1
                
                try:
                    if file_lower.endswith('.xlsb'):
                        df = pd.read_excel(file_path, engine='pyxlsb', header=None)
                    else:
                        df = pd.read_excel(file_path, header=None)
                    
                    # 1. 한국도로공사 건 검사
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
                    except IndexError:
                        pass
                            
                    if not is_dorogongsa:
                        continue
                        
                    dorogongsa_count += 1
                    print(f"✅ 도로공사 건 분석 중 ({dorogongsa_count}): [ {rel_path} ]")
                    
                    # 2. 헤더 행 찾기
                    header_row_idx = -1
                    for idx, row in df.iterrows():
                        if str(row[0]).replace(" ", "") == "순위" and str(row[1]).replace(" ", "") == "회사명":
                            header_row_idx = idx
                            break
                            
                    if header_row_idx == -1: continue
                        
                    headers = [str(x).replace("\n", "").replace(" ", "") for x in df.iloc[header_row_idx].values]
                    
                    def find_col(keywords):
                        for kw in keywords:
                            for c_idx, h in enumerate(headers):
                                if kw in h: return c_idx
                        return -1
                        
                    idx_company = find_col(['회사명'])
                    idx_yega = find_col(['예가대비'])
                    idx_price_score = find_col(['가격점수'])
                    idx_deduct_score = find_col(['단가감점'])
                    
                    if -1 in [idx_company, idx_yega, idx_price_score, idx_deduct_score]: continue
                        
                    data_df = df.iloc[header_row_idx+1:]
                    
                    # 데이터 추출
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
                        winner_company = min_yega_bid['company']
                        
                        results.append({
                            'file': rel_path,
                            'limit_yega': limit_yega,
                            'winner': winner_company
                        })
                        
                        # [핵심] 경쟁사 성향 분석 (해당 공구 기준)
                        for b in valid_bids:
                            comp = b['company']
                            company_stats[comp]['total_bids'] += 1
                            
                            # 해당 업체의 투찰율 - 하한선 차이 
                            # (양수: 안전하게 씀, 0: 하한선 적중, 음수: 한도 초과 공격적/감점)
                            diff = b['yega_ratio'] - limit_yega
                            company_stats[comp]['diffs'].append(diff)
                            
                            if comp == winner_company:
                                company_stats[comp]['wins'] += 1
                                
                            # 공격적 투찰 판별 (가격점수 만점 실패, 보통 하한선 미달로 인한 0.001점 등 감점)
                            if b['price_score'] < max_price_score and b['yega_ratio'] < limit_yega:
                                company_stats[comp]['aggresive_count'] += 1
                            
                            # 보수적 투찰 판별 (안전하게 만점은 받았지만 금액이 하한선 대비 0.5% 이상 높음)
                            if b['price_score'] == max_price_score and diff > 0.5:
                                company_stats[comp]['conservative_count'] += 1

                except Exception as e:
                    print(f"  -> 에러: {e}")
                    
    # 성향 통계 계산 및 정렬
    for comp in company_stats:
        diffs = company_stats[comp]['diffs']
        if diffs:
            company_stats[comp]['avg_diff_from_limit'] = sum(diffs) / len(diffs)
            
    # 전체 하한선 분석 요약
    if results:
        print("\n" + "="*50)
        print(" 📊 전체 한국도로공사 입찰 통계 요약")
        print("="*50)
        ratios = [r['limit_yega'] for r in results]
        print(f"[검색 통계] 총 {scanned_count}개 파일 스캔 중 도로공사 건 {len(results)}건 분석 완료")
        print(f"- 하한선(예가대비) 최소값 : {min(ratios):.3f}%")
        print(f"- 하한선(예가대비) 최대값 : {max(ratios):.3f}%")
        print(f"- 하한선(예가대비) 평균값 : {sum(ratios)/len(ratios):.3f}%")
        
        # 경쟁사 성향 Top 리스트
        # 참여 횟수 3회 이상인 주요 업체만 필터링
        core_companies = {k: v for k, v in company_stats.items() if v['total_bids'] >= 3}
        
        print("\n" + "="*50)
        print(" 🎯 주요 경쟁사 투찰 성향 분석 (참여 3회 이상)")
        print("="*50)
        
        # 정밀 타격 잘하는 업체 (Wins 기준 내림차순)
        sorted_by_wins = sorted(core_companies.items(), key=lambda x: x[1]['wins'], reverse=True)
        print("\n[🥇 정밀 타격 우수 업체 - 하한선 적중 횟수 순]")
        for i, (comp, stat) in enumerate(sorted_by_wins[:5]):
            win_rate = (stat['wins'] / stat['total_bids']) * 100
            print(f" {i+1}. {comp}: {stat['wins']}회 적중 / 총 {stat['total_bids']}회 참여 (승률 {win_rate:.1f}%)")
            
        # 가장 공격적인 업체 (감점 불사하고 낮게 쓰는 성향)
        sorted_by_agg = sorted(core_companies.items(), key=lambda x: x[1]['aggresive_count']/x[1]['total_bids'], reverse=True)
        print("\n[🔥 초공격적 투자 성향 업체 - 감점 감수 하한선 돌파율 순]")
        for i, (comp, stat) in enumerate(sorted_by_agg[:5]):
            agg_rate = (stat['aggresive_count'] / stat['total_bids']) * 100
            print(f" {i+1}. {comp}: 총 참여 {stat['total_bids']}회 중 {stat['aggresive_count']}회 돌파 (돌파율 {agg_rate:.1f}%)")
            
        # 하한선에 가장 근접하게 쓰는 업체 (평균 갭이 0에 가까운 순)
        sorted_by_gap = sorted(core_companies.items(), key=lambda x: abs(x[1]['avg_diff_from_limit']))
        print("\n[📐 하한선 갭 최소화 업체 - 평균 갭 순]")
        for i, (comp, stat) in enumerate(sorted_by_gap[:5]):
            avg_gap = stat['avg_diff_from_limit']
            sign = "+" if avg_gap > 0 else ""
            print(f" {i+1}. {comp} : 평균 갭 {sign}{avg_gap:.3f}% (총 {stat['total_bids']}회 참여)")

if __name__ == "__main__":
    analyze_bids(BASE_DIR)
