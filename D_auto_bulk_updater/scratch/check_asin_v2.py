import pandas as pd
import os

master_path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx'

if os.path.exists(master_path):
    df = pd.read_excel(master_path, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
    
    expr_col = next((c for c in df.columns if "Product Targeting Expression" in str(c)), "Product Targeting Expression")
    camp_id_col = next((c for c in df.columns if "Campaign ID" in str(c)), "Campaign ID")
    camp_name_col = next((c for c in df.columns if "Campaign Name" in str(c)), "Campaign Name")
    
    target_asin = 'B0F139DKR7'
    asin_row = df[df[expr_col].astype(str).str.contains(target_asin, na=False)].iloc[0]
    
    cid = asin_row[camp_id_col]
    print(f"Campaign ID for ASIN {target_asin}: {cid}")
    
    camp_name = df[df[camp_id_col] == cid][camp_name_col].dropna().unique()
    print(f"Campaign Name for this ID: {camp_name}")
