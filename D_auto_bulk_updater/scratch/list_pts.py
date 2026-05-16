import pandas as pd
import os

master_path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx'

if os.path.exists(master_path):
    df = pd.read_excel(master_path, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
    
    camp_name_col = next((c for c in df.columns if "Campaign Name" in str(c)), "Campaign Name")
    camp_id_col = next((c for c in df.columns if "Campaign ID" in str(c)), "Campaign ID")
    
    target_camp = 'B0D3PR4S5B_BB02_PT_(BBJ,PJ,RJ)_221025'
    
    # Find Campaign ID
    cids = df[df[camp_name_col] == target_camp][camp_id_col].unique()
    if len(cids) > 0:
        cid = cids[0]
        print(f"Campaign ID for '{target_camp}': {cid}")
        
        # List all product targeting in this campaign
        pt_rows = df[(df[camp_id_col] == cid) & (df['Entity'].str.lower() == 'product targeting')]
        print(f"Found {len(pt_rows)} product targets:")
        for _, row in pt_rows.iterrows():
            print(f"  - '{row['Product Targeting Expression']}'")
    else:
        print(f"Campaign '{target_camp}' not found in master.")
