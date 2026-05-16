import pandas as pd
import os

master_files = [
    r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx',
    r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 2\BulkSheetExport_3004-0405.xlsx'
]

targets_to_check = [
    ('B0F62HNF16_ ONM_NURSE_KT_nurse', 'nurse gifts'),
    ('B0F62HNF16_ ONM_NURSE_KT_nurse', 'nurse gifts for women'),
    ('B0F62HNF16_ ONM_NURSE_KT_nurse', 'nurse ornament')
]

for master_path in master_files:
    print(f"\nChecking master file: {master_path}")
    if not os.path.exists(master_path):
        print("File not found.")
        continue
        
    try:
        # Load the relevant sheet
        df = pd.read_excel(master_path, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
        
        # Identity columns
        entity_col = next((c for c in df.columns if "Entity" in str(c)), None)
        camp_name_col = next((c for c in df.columns if "Campaign Name" in str(c)), None)
        kw_text_col = next((c for c in df.columns if "Keyword Text" in str(c)), None)
        target_expr_col = next((c for c in df.columns if "Product Targeting Expression" in str(c)), None)
        
        print(f"Columns found: Entity={entity_col}, Campaign Name={camp_name_col}, Keyword Text={kw_text_col}, Target Expression={target_expr_col}")
        
        for camp_name, target in targets_to_check:
            print(f"\nSearching for Campaign: '{camp_name}', Target: '{target}'")
            
            # Filter rows for this campaign
            camp_rows = df[df[camp_name_col].astype(str) == camp_name]
            if camp_rows.empty:
                # Try partial match or stripped match
                camp_rows = df[df[camp_name_col].astype(str).str.strip() == camp_name.strip()]
                if camp_rows.empty:
                    print(f"  -> Campaign not found.")
                    continue
            
            # Find the keyword/target row
            found = False
            for _, row in camp_rows.iterrows():
                row_entity = str(row.get(entity_col, '')).lower()
                row_target = ""
                if 'keyword' in row_entity:
                    row_target = str(row.get(kw_text_col, ''))
                elif 'product targeting' in row_entity:
                    row_target = str(row.get(target_expr_col, ''))
                
                if row_target.strip() == target.strip():
                    print(f"  -> FOUND! Entity: {row_entity}")
                    # Print metrics
                    metrics = ['Impressions', 'Clicks', 'Spend', 'Sales', 'Orders']
                    for m in metrics:
                        val = row.get(m, 'N/A')
                        print(f"     {m}: {val}")
                    found = True
                    break
            
            if not found:
                print(f"  -> Target not found in this campaign.")
                
    except Exception as e:
        print(f"Error reading file: {e}")
