import pandas as pd
import os

master_path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx'

if os.path.exists(master_path):
    df = pd.read_excel(master_path, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
    sku_col = next((c for c in df.columns if "SKU" in str(c)), "SKU")
    entity_col = next((c for c in df.columns if "Entity" in str(c)), "Entity")
    
    # Check row 398
    row_398 = df.iloc[398]
    print(f"Row 398 Entity: {row_398[entity_col]}")
    print(f"Row 398 SKU: {row_398[sku_col]}")
