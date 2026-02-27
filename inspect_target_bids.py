import gspread
from google.oauth2.service_account import Credentials
import os
import pandas as pd
import sys
import traceback

sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = 'credentials.json'
SHEET_ID = '1cENqgfNjt3xZYMGr9LAp71S1b67BS2PRTCXtVHhsudU'
WORKSHEET_NAME = '조달청 300억 이상 토목공사 낙찰현황'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def check_google_sheet():
    print("--- 1. Checking Target Google Sheet Template ---")
    try:
        credentials = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        gc = gspread.authorize(credentials)
        sh = gc.open_by_key(SHEET_ID)
        
        print("Available worksheets:", [ws.title for ws in sh.worksheets()])
        
        try:
            worksheet = sh.worksheet(WORKSHEET_NAME)
        except gspread.WorksheetNotFound:
            print(f"Worksheet '{WORKSHEET_NAME}' not found. Trying the first sheet.")
            worksheet = sh.worksheets()[0]
            
        data = worksheet.get_all_values()
        print(f"Extracted {len(data)} rows from '{worksheet.title}'.")
        for i in range(min(10, len(data))):
            print(f"Row {i+1}:", data[i])
            
    except Exception as e:
        print(f"Error accessing Google Sheet:")
        traceback.print_exc()

def check_local_xlsb():
    print("\n--- 2. Checking Local .xlsb Files ---")
    base_dir = r"E:\인프라수주팀\트레이닝\24년이후 입찰결과"
    
    xlsb_files = []
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.endswith(".xlsb"):
                xlsb_files.append(os.path.join(root, file))
                
    print(f"Found {len(xlsb_files)} .xlsb files.")
    
    if xlsb_files:
        target_file = xlsb_files[0]
        print(f"Inspecting file: {target_file}")
        try:
            df = pd.read_excel(target_file, engine='pyxlsb', header=None)
            print("Dataframe shape:", df.shape)
            print("First 15 rows:")
            print(df.head(15))
        except Exception as e:
            print(f"Error reading xlsb: {e}")

if __name__ == "__main__":
    check_google_sheet()
    check_local_xlsb()
