import pandas as pd
import os

output_file = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 4\PPC_Musemory_UPDATED.xlsx'

xl = pd.ExcelFile(output_file)
print(f"Sheets: {xl.sheet_names}")

for sheet in xl.sheet_names:
    if sheet == 'Listing': continue
    df = pd.read_excel(output_file, sheet_name=sheet, header=None)
    print(f"\nAnalyzing sheet: {sheet}")
    
    # Try to find "Ngày" in the first row
    current_block = None
    max_cols = df.shape[1]
    for start_col in range(6, max_cols, 12):
        cell_val = str(df.iloc[0, start_col])
        if "Ngày" in cell_val:
            current_block = start_col
    
    if current_block is None:
        print(f"No date block found in sheet {sheet}")
        continue
        
    print(f"Latest block starts at column index: {current_block} ({df.iloc[0, current_block]})")
    
    missing_data_rows = []
    for i in range(2, len(df)):
        campaign = df.iloc[i, 1]
        target = df.iloc[i, 3]
        if pd.isna(campaign) or pd.isna(target): continue
        
        metrics = df.iloc[i, current_block:current_block+11]
        
        # Check if all metrics are 0 or NaN
        # Note: Impressions, Clicks, etc. should be numbers
        is_missing = True
        try:
            for val in metrics:
                if pd.notna(val) and float(val) > 0:
                    is_missing = False
                    break
        except:
            pass
            
        if is_missing:
            missing_data_rows.append({
                'Row': i + 1,
                'Campaign': campaign,
                'Target': target,
                'Metrics': metrics.tolist()
            })
            if len(missing_data_rows) >= 5: break

    if missing_data_rows:
        print("Sample rows with missing data:")
        for row in missing_data_rows:
            print(row)
    else:
        print("No missing data found in this sheet (for the latest block).")
