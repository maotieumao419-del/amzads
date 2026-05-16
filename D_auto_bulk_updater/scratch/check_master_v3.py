import pandas as pd
import os

master_path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx'

if os.path.exists(master_path):
    df = pd.read_excel(master_path, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
    
    kw_text_col = next((c for c in df.columns if "Keyword Text" in str(c)), None)
    camp_name_col = next((c for c in df.columns if "Campaign Name" in str(c)), None)
    entity_col = next((c for c in df.columns if "Entity" in str(c)), None)
    state_col = next((c for c in df.columns if "State" in str(c)), "State")
    
    nurse_gifts_rows = df[df[kw_text_col].astype(str).str.contains('nurse gifts', case=False, na=False)]
    
    print(f"Found {len(nurse_gifts_rows)} rows with 'nurse gifts':")
    for idx, row in nurse_gifts_rows.iterrows():
        print(f"Index: {idx}, Campaign: {row[camp_name_col]}, Entity: {row[entity_col]}, State: {row.get(state_col)}")
        # Check parent campaign
        # Usually for a keyword, the campaign name is in the same row or a parent row
        # In Bulksheets 2.0, Keyword rows have Campaign Name?
        # Let's see what else is in this row
        metrics = ['Impressions', 'Clicks', 'Spend', 'Sales', 'Orders']
        m_vals = {m: row.get(m, 0) for m in metrics}
        print(f"  Metrics: {m_vals}")

    # Let's also check for 'nurse gifts for women'
    print("\nChecking for 'nurse gifts for women':")
    women_rows = df[df[kw_text_col].astype(str).str.contains('nurse gifts for women', case=False, na=False)]
    for idx, row in women_rows.iterrows():
        print(f"Index: {idx}, Campaign: {row[camp_name_col]}, Entity: {row[entity_col]}, State: {row.get(state_col)}")
        metrics = ['Impressions', 'Clicks', 'Spend', 'Sales', 'Orders']
        m_vals = {m: row.get(m, 0) for m in metrics}
        print(f"  Metrics: {m_vals}")
