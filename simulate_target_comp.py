import os
import sys
import pandas as pd
import warnings
from collections import defaultdict

# 경고 무시
warnings.filterwarnings('ignore')
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"

def simulate_target_bidding(base_dir):
    # 공구별 하한선 데이터 저장
    bidding_data = [] 
    
    # 업체별 시뮬레이션 결과 저장
    # target_company -> { 'total_encounters': 0, 'success_wins': 0, 'fail_deduction': 0, 'fail_too_high': 0 }
    target_stats = defaultdict(lambda: {
        'total_encounters': 0, 
        'success_wins': 0,      # 타겟사보다 0.001% 낮게 썼을 때 1순위를 달성한 횟수
        'fail_deduction': 0,    # 타겟사보다 낮게 썼다가 하한선 미달로 감점(-0.001점 등) 받은 횟수
        'fail_too_high': 0      # 타겟사보다 낮게 썼지만, 다른 업체가 더 낮게 써서 1위를 뺏긴 횟수
    })
    
    scanned_count = 0
    dorogongsa_count = 0
    
    print("도로공사 엑셀 데이터 파싱 및 하한선 매핑 중...")
    
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
                    
                    # 한국도로공사 검사
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
                    
                    # 헤더 탐색
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
                            'file': rel_path,
                            'limit_yega': limit_yega, # 이 공구의 1순위 만점 하한선
                            'bids': valid_bids        # 전체 참여자 투찰 데이터
                        })

                except Exception as e:
                    print(f"에러: {e}")
                    
    print(f"데이터 파싱 완료 (총 {dorogongsa_count}개 도로공사 건)")
    print("\n[🎯 타겟팅 시뮬레이션 시작]")
    print("조건: 특정 업체(Target)보다 항상 '0.001%' 낮게 투찰했을 때의 결과 추적\n")
    
    # 2. 타겟팅 시뮬레이션
    target_margin = 0.001 # 타겟보다 0.001% 미세하게 낮게 투찰한다는 가상의 전략
    
    for data in bidding_data:
        limit_yega = data['limit_yega']
        bids = data['bids']
        
        # 이번 입찰에 참여한 모든 업체들을 순회하며, "이 업체를 타겟으로 찍었다면?" 시뮬레이션
        for target_bid in bids:
            target_comp = target_bid['company']
            my_sim_yega = target_bid['yega_ratio'] - target_margin # 타겟보다 아슬아슬하게 낮게!
            
            target_stats[target_comp]['total_encounters'] += 1
            
            # 나의 투찰 결과 판별
            if my_sim_yega < limit_yega:
                # [실패1] 타겟보다 낮게 썼다가, 투찰 하한선을 깨버려서 가격점수에서 감점을 당함
                target_stats[target_comp]['fail_deduction'] += 1
            else:
                # 감점은 면함. 그렇다면 1위를 했을까? (나보다 더 낮게, 감점 없이 쓴 놈이 있는가?)
                # limit_yega <= 다른업체들_yega <= my_sim_yega 인 업체가 있는지 확인
                is_first = True
                for other_bid in perfect_bids: # 이 공구에서 만점/무감점 받은 업체들
                    if limit_yega <= other_bid['yega_ratio'] < my_sim_yega:
                        is_first = False
                        break
                
                if is_first:
                    # [성공] 감점을 받지 않는 선에서 가장 낮게 투찰하여 1위(우선순위) 달성 완료!
                    target_stats[target_comp]['success_wins'] += 1
                else:
                    # [실패2] 하한선 안에는 들어와서 만점은 받았지만, 더 낮게 쓴 타사에게 1위를 뺏김 (보수적 투찰)
                    target_stats[target_comp]['fail_too_high'] += 1

    # 3. 시뮬레이션 결과 집계 및 출력 (최소 5번 이상 등장한 업체 대상 계산)
    valid_targets = {k: v for k, v in target_stats.items() if v['total_encounters'] >= 3}
    
    # 승률 높은 순 정렬
    sorted_targets = sorted(
        valid_targets.items(), 
        key=lambda x: (x[1]['success_wins'] / x[1]['total_encounters']), 
        reverse=True
    )
    
    print("="*70)
    print(" 💡 시뮬레이션 결과: '누구보다 조금 덜 벌겠다(-0.001%)'고 했을 때 가장 안전하고 승률 높은 타겟 업체 TOP 10")
    print("="*70)
    
    for i, (comp, stats) in enumerate(sorted_targets[:10]):
        total = stats['total_encounters']
        wins = stats['success_wins']
        fail_deduct = stats['fail_deduction']
        fail_high = stats['fail_too_high']
        
        win_rate = (wins / total) * 100
        deduct_rate = (fail_deduct / total) * 100
        
        print(f"{i+1}위 타겟: [{comp}]")
        print(f"   ▶ {comp}만 따라다니며 {target_margin}% 낮게 투찰할 경우: 시뮬레이션 승률 {win_rate:.1f}%")
        print(f"      (총 조우 {total}회 중 1순위 낙찰 {wins}회 / 감점탈락 {fail_deduct}회 / 보수적탈락 {fail_high}회)")
        
        # 전략적 해석 브리핑
        if win_rate > 30:
            msg = f"      *분석: {comp}는 평소 하한선에 아주 근접하는 타겟입니다. 이 업체만 이긴다는 마인드로 0.001% 낮게 써도 감점(하한선 뚫림)을 당할 확률이 {deduct_rate:.1f}%밖에 되지 않아 훌륭한 벤치마크 타겟입니다."
        elif deduct_rate > 50:
            msg = f"      *위험: {comp} 자체도 워낙 하한선 뚫고 출혈 경쟁을 하는 타입이라, 이 업체보다 낮게 쓰면 무려 {deduct_rate:.1f}% 확률로 동반 감점 자폭을 하게 됩니다. 거르세요!"
        else:
            msg = f"      *무난: {comp}는 지나치게 높게 씁니다. 만약 이 업체보다 조금 더 낮게 쓰더라도, 더 싼 회사가 나타나서 낙찰받을 확률이 높습니다."
        print(msg)
        print("-" * 70)

if __name__ == "__main__":
    simulate_target_bidding(BASE_DIR)
