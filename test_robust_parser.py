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

for f in files:
    try:
        df = pd.read_excel(f, sheet_name="입찰결과", header=None, engine='pyxlsb' if f.endswith('xlsb') else 'openpyxl')
        for i, row in df.iterrows():
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
                    # We found a company in this row. Now let's extract ratio and bid amount from this row.
                    ratio = None
                    bid = None
                    for c in row_vals:
                        if pd.isna(c): continue
                        
                        # Try to parse as float
                        try:
                            # if it's a string, it might have % or commas
                            c_str = str(c).replace(',', '').replace('%', '').strip()
                            val = float(c_str)
                            
                            # if value is between 0.75 and 0.99, it's likely the ratio
                            if 0.75 <= val <= 0.99:
                                ratio = val
                            # sometimes it's written as 75.0 ~ 99.0
                            elif 75.0 <= val <= 99.0:
                                ratio = val / 100.0
                            # if value is very large, it's the bid amount
                            elif val > 100000000:
                                bid = val
                        except ValueError:
                            pass
                            
                    all_data.append({
                        'File': os.path.basename(f),
                        'Company': matched_comp,
                        'Faction': matched_faction,
                        'Ratio': ratio,
                        'BidAmount': bid
                    })
                    break # move to next row to avoid double counting same row
    except Exception as e:
        print(f"Error reading {f}: {e}")

res_df = pd.DataFrame(all_data)
# Filter valid ratios
res_df = res_df.dropna(subset=['Ratio'])
print(f"Total valid unique files processed: {res_df['File'].nunique()} / {len(files)}")
print(f"Total rows extracted: {len(res_df)}")
