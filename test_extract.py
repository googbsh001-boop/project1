import pandas as pd
import os

folder_path = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"
target_file = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과\입찰결과 - 260113 (종심-고) 새만금항 신항 방파제(연장) 축조공사_Rev.N4.10.xlsb"

def format_excel_date(excel_date):
    if pd.isna(excel_date): return ""
    try:
        # Excel date leap year bug adjustment defaults to 1899-12-30
        dt = pd.to_datetime('1899-12-30') + pd.to_timedelta(float(excel_date), unit='D')
        if dt.hour == 0 and dt.minute == 0:
            return dt.strftime('%Y/%m/%d')
        else:
            return dt.strftime('%Y/%m/%d %H:%M')
    except:
        return str(excel_date)

def test_extract():
    df = pd.read_excel(target_file, engine='pyxlsb', header=None)
    
    project_name = ""
    client_str = ""
    bid_date_str = ""
    header_row_idx = -1
    
    for idx, row in df.head(30).iterrows():
        c0 = str(row[0]).replace(" ", "")
        if c0 == "공사명":
            project_name = str(row[2]).strip() if not pd.isna(row[2]) else ""
        elif c0 == "발주처":
            client_str = str(row[2]).strip() if not pd.isna(row[2]) else ""
        elif c0 == "입찰일":
            bid_date_str = format_excel_date(row[2])
        elif c0 == "순위":
            header_row_idx = idx
            break

    print(f"Project: {project_name}")
    print(f"Client: {client_str}")
    print(f"Date: {bid_date_str}")
    
    if header_row_idx != -1:
        data_df = df.iloc[header_row_idx+1:].copy()
        for _, row in data_df.iterrows():
            rank = row[0]
            if pd.isna(rank): continue
            rank_str = str(rank).strip()
            if rank_str == "1" or rank_str == "1.0":
                company = str(row[1]).strip()
                amount = float(row[2])
                try:
                    r = row[4]
                    ratio = float(r) * 100 if not pd.isna(r) else 0
                except:
                    ratio = 0
                print(f"Winner: {company}, Amount: {amount}, Ratio: {round(ratio, 4)}")
                break

if __name__ == "__main__":
    test_extract()
