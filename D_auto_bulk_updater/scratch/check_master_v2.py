import pandas as pd
import os

master_path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx'

if os.path.exists(master_path):
    df = pd.read_excel(master_path, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
    
    camp_name_col = next((c for c in df.columns if "Campaign Name" in str(c)), None)
    kw_text_col = next((c for c in df.columns if "Keyword Text" in str(c)), None)
    
    print("Listing unique campaigns containing 'NURSE':")
    nurses_camps = [c for c in df[camp_name_col].dropna().unique() if 'NURSE' in str(c)]
    for c in nurses_camps:
        print(f"'{c}'")
        
    print("\nChecking if 'nurse gifts' exists anywhere:")
    nurse_gifts_rows = df[df[kw_text_col].astype(str).str.contains('nurse gifts', case=False, na=False)]
    if not nurse_gifts_rows.empty:
        print(f"Found {len(nurse_gifts_rows)} rows. Sample campaigns:")
        print(nurse_gifts_rows[camp_name_col].unique()[:10])
    else:
        print("Not found.")
