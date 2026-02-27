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
            # Find the .BID file in this folder
            bid_file = None
            for f in os.listdir(folder_path):
                if f.upper().endswith('.BID'):
                    bid_file = f
                    break
            
            if bid_file:
                # The first string before space is the index, e.g. "00001"
                parts = folder_name.split()
                if parts:
                    idx = parts[0]
                    full_path = os.path.join(folder_path, bid_file)
                    mapping[idx] = full_path
    return mapping

def main():
    credentials = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.worksheet(WORKSHEET_NAME)
    
    mapping = map_folders()
    print(f"Mapped {len(mapping)} folders.")
    
    # We need to read column A and update column F
    # Let's get all values of Column A
    col_a = worksheet.col_values(1)  # 1-indexed for Col A
    
    # Prepare the list of values to update for Col F
    # Start from row 6? Let's check where '00001' starts.
    # We will just map every row where Col A has a mapping.
    # It's safer to get all of Col F as well to keep existing if no map
    col_f = worksheet.col_values(6)
    
    # Extend col_f to match col_a length if needed
    while len(col_f) < len(col_a):
        col_f.append("")
        
    updates_count = 0
    for i, a_val in enumerate(col_a):
        idx = a_val.strip()
        if idx in mapping:
            col_f[i] = mapping[idx]
            updates_count += 1
            
    if updates_count > 0:
        # worksheet.update requires a list of lists
        update_data = [[val] for val in col_f]
        
        # update from F1 to F...
        cell_range = f'F1:F{len(update_data)}'
        worksheet.update(values=update_data, range_name=cell_range)
        print(f"Successfully updated {updates_count} cells in Column F.")
    else:
        print("No matches found to update.")

if __name__ == "__main__":
    main()
