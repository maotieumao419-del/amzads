import pandas as pd
import os

master_path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx'

if os.path.exists(master_path):
    df = pd.read_excel(master_path, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
    
    camp_id_col = next((c for c in df.columns if "Campaign ID" in str(c)), None)
    camp_name_col = next((c for c in df.columns if "Campaign Name" in str(c)), None)
    kw_text_col = next((c for c in df.columns if "Keyword Text" in str(c)), None)
    
    # Check rows around 399-403
    print("Rows around 399-403:")
    sample = df.iloc[395:410][ [camp_id_col, camp_name_col, kw_text_col, 'Entity'] ]
    print(sample)
    
    # Let's find the campaign name for the campaign ID of row 399
    cid = df.iloc[399][camp_id_col]
    print(f"\nCampaign ID for row 399: {cid}")
    
    camp_info = df[df[camp_id_col] == cid]
    print(f"Number of rows in this campaign group: {len(camp_info)}")
    c_names = camp_info[camp_name_col].dropna().unique()
    print(f"Campaign Names found in this group: {c_names}")
