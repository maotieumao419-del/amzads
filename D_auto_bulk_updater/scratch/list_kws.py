import pandas as pd
import os

master_path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx'

if os.path.exists(master_path):
    df = pd.read_excel(master_path, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
    
    camp_name_col = next((c for c in df.columns if "Campaign Name" in str(c)), "Campaign Name")
    camp_id_col = next((c for c in df.columns if "Campaign ID" in str(c)), "Campaign ID")
    
    target_camp = 'B0F62HNF16_ ONM_NURSE_KT_nurse'
    
    # Find Campaign ID for this name
    cid = df[df[camp_name_col] == target_camp][camp_id_col].unique()
    if len(cid) > 0:
        cid = cid[0]
        print(f"Campaign ID for '{target_camp}': {cid}")
        
        # List all keywords in this campaign
        kw_rows = df[(df[camp_id_col] == cid) & (df['Entity'].str.lower() == 'keyword')]
        print(f"Found {len(kw_rows)} keywords:")
        for _, row in kw_rows.iterrows():
            print(f"  - '{row['Keyword Text']}' (Match Type: {row['Match Type']})")
    else:
        print(f"Campaign '{target_camp}' not found in master.")
