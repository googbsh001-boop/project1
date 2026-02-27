import gspread
from google.oauth2.service_account import Credentials
import os
import pandas as pd
import sys
import re
from gspread_formatting import format_cell_range, CellFormat, TextFormat, Color, NumberFormat, set_column_width, Borders, Border

sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
BASE_DIR = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"
SHEET_ID = "1n3WxFMxjS-mhHGE8I4dXi4Q2oJ3l4sq_OkCkBeJkbJI"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def format_excel_date(excel_date):
    if pd.isna(excel_date) or str(excel_date).strip() == "": return ""
    try:
        dt = pd.to_datetime('1899-12-30') + pd.to_timedelta(float(excel_date), unit='D')
        if dt.hour == 0 and dt.minute == 0:
            return dt.strftime('%Y/%m/%d')
        else:
            return dt.strftime('%Y/%m/%d %H:%M')
    except:
        return str(excel_date)

def get_bid_date(file_path):
    try:
        df_info = pd.read_excel(file_path, engine='pyxlsb', sheet_name='기초정보', header=None)
        
        r1_val = None
        try:
            r1_val = df_info.iloc[0, 17]
        except:
            pass
            
        date_str = format_excel_date(r1_val)
        if date_str:
            return date_str
            
        for idx, row in df_info.iterrows():
            for col_idx, cell in enumerate(row):
                if str(cell).replace(" ", "") == "입찰마감":
                    next_val = row[col_idx + 1] if col_idx + 1 < len(row) else None
                    if next_val and not pd.isna(next_val):
                        return format_excel_date(next_val)
                    
        return ""
    except Exception as e:
        print(f"[{os.path.basename(file_path)}] Failed to read 기초정보: {e}")
        return ""

def process_file_rank1(file_path):
    try:
        filename = os.path.basename(file_path)
        
        # 1. Fetch info from "기초정보" sheet
        bid_date_str = get_bid_date(file_path)
            
        # 2. Main sheet parsing
        main_sheet_name = None
        xls = pd.ExcelFile(file_path, engine='pyxlsb')
        for sn in xls.sheet_names:
            if "기초" not in sn:
                main_sheet_name = sn
                break
        
        if not main_sheet_name:
            main_sheet_name = xls.sheet_names[0]
            
        df = pd.read_excel(file_path, engine='pyxlsb', sheet_name=main_sheet_name, header=None)
        
        decision_method = ""
        try:
            val_v1 = df.iloc[0, 21]
            if not pd.isna(val_v1):
                decision_method = str(val_v1).strip()
        except:
            pass
        
        project_name = ""
        client_str = ""
        base_amount = 0 # 기초금액
        balance_price = 0 # 균형가격
        header_row_idx = -1
        
        for idx, row in df.head(30).iterrows():
            c0 = str(row[0]).replace(" ", "")
            if c0 == "공사명":
                project_name = str(row[2]).strip() if not pd.isna(row[2]) else ""
            elif c0 == "발주처":
                client_str = str(row[2]).strip() if not pd.isna(row[2]) else ""
            elif c0 == "순위":
                header_row_idx = idx
                break
                
            # Search for base_amount (기초금액) and balance_price (균형가격)
            for c_idx, cell in enumerate(row):
                cell_str = str(cell).replace(" ", "")
                if "기초금액" in cell_str:
                    try:
                        # usually in I4. In our loop, we look for the next numeric value in the same row
                        # but in the user test it was exactly at I4 and value was at I4. Let's just grab row[8] (index 8 is I)
                        # Wait, the label "기초금액" was at G4 (col 6), and value was at I4 (col 8).
                        # Let's just robustly grab I4 and I6 directly by index.
                        pass
                    except:
                        pass
        
        try:
            # I4 is index 3, 8
            # I6 is index 5, 8
            base_val = df.iloc[3, 8]
            bal_val = df.iloc[5, 8]
            if not pd.isna(base_val): base_amount = float(base_val)
            if not pd.isna(bal_val): balance_price = float(bal_val)
        except Exception as e:
            print(f"Error extracting I4/I6 for {filename}: {e}")
        
        if header_row_idx == -1: return None

        data_df = df.iloc[header_row_idx+1:].copy()
        
        note_str = ""
        match_note = re.search(r'\((.*?)\)', filename)
        if match_note:
            note_str = match_note.group(1).strip()
            
        for _, row in data_df.iterrows():
            try:
                if len(row) <= 13: continue 
                n_col = row[13] 
                if pd.isna(n_col): continue
                
                n_str = str(n_col).strip()
                
                if n_str == "1" or n_str == "1.0":
                    company = str(row[1]).strip()
                    amount = float(row[2]) 
                    
                    ratio = "-"
                    try:
                        r = row[4] 
                        if not pd.isna(r):
                            ratio_val = float(r) * 100
                            ratio = round(ratio_val, 4)
                    except:
                        pass
                    
                    if not project_name:
                        match = re.search(r'\)(.*?)(?:_Rev|\.xlsb)', filename)
                        if match:
                            project_name = match.group(1).strip()
                        else:
                            project_name = filename.replace('.xlsb', '').strip()
                            
                    clean_client = client_str
                    if "조달청(" in client_str and client_str.endswith(")"):
                        clean_client = client_str.replace("조달청(", "").rstrip(")")
                        
                    return {
                        'date': bid_date_str,
                        'project': project_name,
                        'client': clean_client,
                        'company': company,
                        'decision_method': decision_method,
                        'amount': int(amount),
                        'base_amount': int(base_amount) if base_amount else 0,
                        'balance_price': int(balance_price) if balance_price else 0,
                        'ratio': ratio,
                        'note': note_str
                    }
            except Exception as e:
                continue
                
    except Exception as e:
        print(f"Error parsing {file_path}: {e}")
        return None
    return None

def main():
    print("--- Parsing local .xlsb files for Rank 1 Winner Data & Add Base/Balance ---")
    xlsb_files = []
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if file.startswith("입찰결과") and file.endswith(".xlsb"):
                xlsb_files.append(os.path.join(root, file))
                
    results = []
    for file_path in xlsb_files:
        data = process_file_rank1(file_path)
        if data:
            results.append(data)
            
    results.sort(key=lambda x: (x['date'], x['project']))
    
    # Header definition
    # The user asked:
    # "D"열 오른쪽에 "기초금액","균형가격", "균형/기초", "투찰/기초"열을 각각 추가삽입
    # Current headers up to D: 입찰일시(A) | 공고명(B) | 수요기관(C) | 낙찰자(D)
    # New: 기초금액(E) | 균형가격(F) | 균형/기초(G) | 투찰/기초(H)
    # Afterwards, we need 투찰금액(I) and 비고(J)? Wait, the user said "E1, E2 칸은 투찰금액(원), 기초대비(%) 로 입력" previously.
    # So if we INSERT 4 columns to the right of D... then E, F, G, H will be the new columns.
    # And I, J will be "투찰금액", "투찰률". OR the user might want D, E, F, G, H, I, J, K.
    # Let's map exactly as requested:
    # A: 입찰일시
    # B: 공고명
    # C: 수요기관
    # D: 낙찰자
    # E: 기초금액
    # F: 균형가격
    # G: 균형/기초 (%)
    # H: 투찰/기초 (%) = (투찰금액 / 기초금액) * 100
    # I: 투찰금액(원)
    # J: 비고
    
    upload_rows = [["입찰일시", "공고명", "수요기관", "낙찰자", "결정방식", "기초금액(원)", "균형가격(원)", "균형/기초(%)", "투찰/기초(%)", "투찰금액(원)", "비고"]]
    for r in results:
        bal_base_ratio = 0
        if r['base_amount'] > 0:
            bal_base_ratio = (r['balance_price'] / r['base_amount']) * 100
            
        bid_base_ratio = 0
        if r['base_amount'] > 0:
            bid_base_ratio = (r['amount'] / r['base_amount']) * 100
            
        upload_rows.append([
            r['date'],
            r['project'],
            r['client'],
            r['company'],
            r['decision_method'],
            r['base_amount'],
            r['balance_price'],
            round(bal_base_ratio, 4),
            round(bid_base_ratio, 4),
            r['amount'],
            r['note']
        ])
        
    print(f"Extracted {len(results)} valid rows.")
    
    print("\n--- Connecting to User's Google Sheet ---")
    credentials = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    gc = gspread.authorize(credentials)
    
    try:
        sh = gc.open_by_key(SHEET_ID)
        worksheet = sh.sheet1
        
        print("Uploading new data format to sheet...")
        worksheet.clear()
        worksheet.update(upload_rows)
        
        print("Applying formatting rules...")
        total_rows = len(upload_rows)
        
        header_fmt = CellFormat(
            backgroundColor=Color(0.2, 0.46, 0.33),
            textFormat=TextFormat(bold=True, foregroundColor=Color(1, 1, 1))
        )
        format_cell_range(worksheet, "A1:K1", header_fmt)
        
        # Money format for F, G, J (기초금액, 균형가격, 투찰금액)
        num_fmt = CellFormat(numberFormat=NumberFormat(type='NUMBER', pattern='#,##0'))
        format_cell_range(worksheet, f"F2:G{total_rows}", num_fmt)
        format_cell_range(worksheet, f"J2:J{total_rows}", num_fmt)
        
        # Ratio format for H, I (균형/기초, 투찰/기초)
        ratio_fmt = CellFormat(numberFormat=NumberFormat(type='NUMBER', pattern='0.0###'))
        format_cell_range(worksheet, f"H2:I{total_rows}", ratio_fmt)
        
        # Borders
        border_style = Border(style='SOLID', color=Color(0.8, 0.8, 0.8))
        full_borders = Borders(top=border_style, bottom=border_style, left=border_style, right=border_style)
        border_fmt = CellFormat(borders=full_borders)
        format_cell_range(worksheet, f"A1:K{total_rows}", border_fmt)
        
        # Column Widths
        set_column_width(worksheet, 'A', 130) # Date
        set_column_width(worksheet, 'B', 280) # Project
        set_column_width(worksheet, 'C', 180) # Client
        set_column_width(worksheet, 'D', 150) # Company
        set_column_width(worksheet, 'E', 100) # Decision Method
        set_column_width(worksheet, 'F', 130) # Base Amount
        set_column_width(worksheet, 'G', 130) # Balance Price
        set_column_width(worksheet, 'H', 100) # Bal/Base
        set_column_width(worksheet, 'I', 100) # Bid/Base
        set_column_width(worksheet, 'J', 130) # Bid Amount
        set_column_width(worksheet, 'K', 120) # Note
        
        print("Data upload and formatting complete!")
        print(f"Access point: {sh.url}")
    except Exception as e:
        print(f"Error accessing Google Sheet: {e}")

if __name__ == "__main__":
    main()
