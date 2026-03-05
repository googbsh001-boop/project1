import os
import glob
import pandas as pd
import numpy as np

FACTIONS = {
    '경남기업': '우리(주황)', '극동건설': '우리(주황)', '남광토건': '우리(주황)',
    '삼환기업': '우리(주황)', '쌍용건설': '우리(주황)', '에이치엘디앤아이한라': '우리(주황)', '호반산업': '우리(주황)',
    '계룡건설산업': '개(그린)', '동양건설산업': '개(그린)', '디엘이앤씨': '개(그린)',
    '케이씨씨건설': '개(그린)', '케이알산업': '개(그린)', '코오롱글로벌': '개(그린)',
    '태영건설': '개(그린)', '현대건설': '개(그린)',
    '금호건설': '원숭이(하늘)', '대우건설': '원숭이(하늘)', '동부건설': '원숭이(하늘)',
    '두산건설': '원숭이(하늘)', '롯데건설': '원숭이(하늘)', '비에스한양': '원숭이(하늘)',
    '한양': '원숭이(하늘)', '에이치제이중공업': '원숭이(하늘)', '지에스건설': '원숭이(하늘)',
    '대보건설': '무소속(흰색)', '디엘건설': '무소속(흰색)', '에이치디씨현대산업개발': '무소속(흰색)', '한화': '무소속(흰색)',
}

folder = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"
all_files = glob.glob(os.path.join(folder, "**", "*.xlsb"), recursive=True) + glob.glob(os.path.join(folder, "**", "*.xlsx"), recursive=True)
files = [f for f in all_files if not os.path.basename(f).startswith('~$')]

all_data = []
winners_data = []

for f in files:
    try:
        df = pd.read_excel(f, sheet_name="입찰결과", header=None, engine='pyxlsb' if f.endswith('xlsb') else 'openpyxl')
        
        # 1. 승자 찾기 ("N"열 우선순위가 "1"인 회사, N열 = 인덱스 13, 회사 = 인덱스 1)
        winner_comp = None
        winner_faction = None
        
        for i, row in df.iterrows():
            row_vals = row.values
            if len(row_vals) > 13:
                try:
                    col_n = str(row_vals[13]).strip().replace('.0', '')
                    if col_n == '1' or col_n == '1순위':
                        raw_comp_name = str(row_vals[1]).strip()
                        # 세력 매칭 확인
                        for key, faction in FACTIONS.items():
                            if key in raw_comp_name.replace(' ', ''):
                                winner_comp = key
                                winner_faction = faction
                                break
                        
                        # 세력에 속하지 않더라도 원본 이름 추출
                        if not winner_comp and len(raw_comp_name) >= 2:
                            winner_comp = raw_comp_name
                            # 이름 정리 (주식회사 등 제거)
                            winner_comp = winner_comp.replace('주식회사', '').replace('(주)', '').replace('㈜', '').strip()
                            winner_faction = '기타(세력외)'
                            
                        if winner_comp:
                            winners_data.append({
                                'File': os.path.basename(f),
                                'WinnerCompany': winner_comp,
                                'WinnerFaction': winner_faction if winner_faction else '기타(세력외)'
                            })
                            break
                except Exception:
                    pass

        if not winner_comp:
             winners_data.append({
                'File': os.path.basename(f),
                'WinnerCompany': '미상/확인불가',
                'WinnerFaction': '기타(세력외)'
            })


        # 2. 모든 투찰 데이터 파싱
        for _, row in df.iterrows():
            row_vals = row.values
            for cell in row_vals:
                if pd.isna(cell): continue
                cell_str = str(cell).replace(' ', '')
                matched_comp = None
                matched_faction = None
                for key, faction in FACTIONS.items():
                    if key in cell_str:
                        matched_comp = key
                        matched_faction = faction
                        break
                        
                if matched_comp:
                    ratio = None
                    bid = None
                    for c in row_vals:
                        if pd.isna(c): continue
                        try:
                            # Parse floats defensively
                            c_str = str(c).replace(',', '').replace('%', '').strip()
                            val = float(c_str)
                            if 0.75 <= val <= 0.99:
                                ratio = val
                            elif 75.0 <= val <= 99.0:
                                ratio = val / 100.0
                            elif val > 100000000:
                                bid = val
                        except ValueError:
                            pass
                            
                    all_data.append({
                        'File': os.path.basename(f),
                        'Company': matched_comp,
                        'Faction': matched_faction,
                        'BidAmount': bid,
                        'Ratio': ratio
                    })
                    break 
    except Exception as e:
        print(f"Skipping {f} (Error: {e})")

res_df = pd.DataFrame(all_data)
# Filter valid ratios
res_df = res_df.dropna(subset=['Ratio'])
res_df.to_csv(r"E:\인프라수주팀\트레이닝\프로젝트1\bidding_analysis.csv", index=False, encoding='utf-8-sig')

winners_df = pd.DataFrame(winners_data)
winners_df.to_csv(r"E:\인프라수주팀\트레이닝\프로젝트1\bidding_winners.csv", index=False, encoding='utf-8-sig')

print(f"Processed {len(res_df)} rows of data from {res_df['File'].nunique()} unique valid files.")
print(f"Winners extracted for {len(winners_df)} files.")

# Also calculate summary
if not res_df.empty:
    summary = res_df.groupby(['Faction', 'Company'])['Ratio'].agg(['mean', 'count', 'min', 'max']).reset_index()
    summary = summary.sort_values(by=['Faction', 'mean'])
    summary.to_csv(r"E:\인프라수주팀\트레이닝\프로젝트1\bidding_summary.csv", index=False, encoding='utf-8-sig')
    print("Summary created.")
