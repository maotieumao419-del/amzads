import pandas as pd
import os

master_path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx'

if os.path.exists(master_path):
    df = pd.read_excel(master_path, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
    
    expr_col = next((c for c in df.columns if "Product Targeting Expression" in str(c)), "Product Targeting Expression")
    camp_name_col = next((c for c in df.columns if "Campaign Name" in str(c)), "Campaign Name")
    
    target_asin = 'B0F139DKR7'
    rows = df[df[expr_col].astype(str).str.contains(target_asin, na=False)]
    
    if not rows.empty:
        print(f"Found {len(rows)} rows for ASIN {target_asin}:")
        for idx, row in rows.iterrows():
            print(f"  Campaign: {row[camp_name_col]}, Entity: {row['Entity']}, Expression: {row[expr_col]}")
    else:
        print(f"ASIN {target_asin} not found in master.")
