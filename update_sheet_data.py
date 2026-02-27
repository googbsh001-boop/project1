import gspread
from google.oauth2.service_account import Credentials
import os

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1ICIoAj-KU-BeYKFEvGNXtpCxpUOXa-JS4Kg0Zy_UOFw'
WORKSHEET_NAME = '리스트'
TARGET_DIR = r'C:\Users\00006050\Desktop\연수금곡\1안'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def map_folders():
    mapping = {}
    if not os.path.exists(TARGET_DIR):
        print("Target dir not found")
        return mapping
    
    for folder_name in os.listdir(TARGET_DIR):
        folder_path = os.path.join(TARGET_DIR, folder_name)
        if os.path.isdir(folder_path):
            bid_file = None
            for f in os.listdir(folder_path):
                if f.upper().endswith('.BID'):
                    bid_file = f
                    break
            
            if bid_file:
                # e.g., "00001  53,319,714,588"
                parts = folder_name.split()
                if len(parts) >= 2:
                    idx = parts[0]
                    # amount usually includes commas
                    amount = parts[1]
                    full_path = os.path.join(folder_path, bid_file)
                    mapping[idx] = {'path': full_path, 'amount': amount}
    return mapping

def main():
    credentials = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.worksheet(WORKSHEET_NAME)
    
    mapping = map_folders()
    print(f"Mapped {len(mapping)} folders.")
    
    # We need to read column A (Index). Data usually starts from row 6.
    all_values = dict()
    data = worksheet.get_all_values()
    
    # We will clear the rows that don't match our index.
    # To do this safely and maintain formulas, we will update D and F for mapped rows,
    # and clear the contents of unmapped rows sequentially.
    
    col_d_updates = []
    col_f_updates = []
    
    delete_rows = [] # collect row numbers (1-indexed) to clear or delete
    
    # Find headers / what row starts data
    # We assume '00001' starts at some row, likely 6.
    start_row = 6
    
    for row_idx in range(start_row - 1, len(data)):
        row_num = row_idx + 1
        row = data[row_idx]
        if not row:
            continue
            
        a_val = row[0].strip()
        if a_val in mapping:
            # We have a match for this row
            amount = mapping[a_val]['amount']
            path = mapping[a_val]['path']
            
            col_d_updates.append({'range': f'D{row_num}', 'values': [[amount]]})
            col_f_updates.append({'range': f'F{row_num}', 'values': [[path]]})
        else:
            # Unmapped row: user wants to delete these.
            # We'll clear them first, or maybe delete them.
            if a_val: # Contains something in A but not mapped, or empty
                delete_rows.append(row_num)

    # Batch update for D and F
    if col_d_updates or col_f_updates:
        # worksheet.batch_update requires a list of dicts with 'range' and 'values'
        updates = col_d_updates + col_f_updates
        worksheet.batch_update(updates)
        print(f"Updated {len(col_d_updates)} amounts and {len(col_f_updates)} paths.")
        
    # Delete unused rows in reverse order to not mess up indices
    if delete_rows:
        print(f"Deleting {len(delete_rows)} unmapped rows.")
        for row_num in sorted(delete_rows, reverse=True):
            # For safety, let's just clear the row content instead of actually deleting the row structure,
            # or actually delete. The user said "다 삭제해줘" (delete all).
            worksheet.delete_rows(row_num)
        print("Deletion complete.")

if __name__ == "__main__":
    main()
