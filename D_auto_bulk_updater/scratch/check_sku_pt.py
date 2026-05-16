import pandas as pd
import os

master_path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx'

if os.path.exists(master_path):
    df = pd.read_excel(master_path, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
    
    camp_name = 'B0D3PR4S5B_BB02_PT_(BBJ,PJ,RJ)_221025'
    sku_col = next((c for c in df.columns if "SKU" in str(c)), "SKU")
    
    skus = df[df['Campaign Name'].astype(str) == camp_name][sku_col].dropna().unique()
    if len(skus) == 0:
        # Check Product Ad rows in this campaign group
        camp_id = df[df['Campaign Name'].astype(str) == camp_name]['Campaign ID'].iloc[0]
        skus = df[(df['Campaign ID'] == camp_id) & (df['Entity'].str.lower() == 'product ad')][sku_col].dropna().unique()
        
    print(f"SKUs for campaign '{camp_name}': {skus}")
