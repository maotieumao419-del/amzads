import os
import re
import pandas as pd
import openpyxl
import logging

# Configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_1_DIR = os.path.join(BASE_DIR, "data", "input 1")
INPUT_2_DIR = os.path.join(BASE_DIR, "data", "input 2")

os.makedirs(INPUT_1_DIR, exist_ok=True)
os.makedirs(INPUT_2_DIR, exist_ok=True)

def get_date_prefix(filename):
    """Extract date range (e.g. 20260415-20260416) from filename."""
    match = re.search(r'(\d{8}-\d{8})', filename)
    if match:
        return match.group(1)
    # Generic match for any 8-digit sequence pairs
    digits = re.findall(r'\d{8}', filename)
    if len(digits) >= 2:
        return f"{digits[0]}-{digits[1]}"
    return "UnknownDate"

def main():
    print("=" * 70)
    print(" >>> PPC DATA SYNC: DICTIONARY MAPPING MODE <<<")
    print("=" * 70)

    master_path = os.path.join(INPUT_1_DIR, "PPC_NGUYÊN.xlsx")
    output_path = os.path.join(INPUT_2_DIR, "PPC_NGUYEN_UPDATED.xlsx")

    if not os.path.exists(master_path):
        print(f"--- Error: Master file not found at {master_path}")
        return

    # STEP 1: EXTRACT TARGET LIST
    print("\n[STEP 1] Extracting Target List from Listing sheet...")
    try:
        df_listing = pd.read_excel(master_path, sheet_name='Listing', dtype=str)
        df_listing.columns = [c.strip() for c in df_listing.columns]
        
        required_cols = ['Trạng thái', 'SKU', 'Store']
        if not all(col in df_listing.columns for col in required_cols):
            print(f"--- Error: Sheet Listing is missing columns: {required_cols}")
            return
            
        df_listing['status_clean'] = df_listing['Trạng thái'].astype(str).str.lower().str.strip()
        df_valid = df_listing[df_listing['status_clean'] == 'đã tạo campaign']
        
        target_list = []
        for _, row in df_valid.iterrows():
            target_list.append({
                'SKU': str(row['SKU']).strip(),
                'Store': str(row['Store']).strip().upper(),
                'Portfolio ID': str(row.get('Portfolio ID', row.get('Portfolio Id', ''))).strip()
            })
        print(f"--- Total targets found: {len(target_list)}")
    except Exception as e:
        print(f"--- Error reading Listing: {e}")
        return

    # STEP 2: PRE-PROCESS BULK DATA (DICTIONARY MAPPING)
    print("\n[STEP 2] Pre-processing Bulk Files into Dictionary...")
    store_data_frames = {} # {STORE_NAME: {'df': df_agg, 'date': date_str}}
    
    try:
        all_files = os.listdir(INPUT_1_DIR)
        for fname in all_files:
            if not fname.endswith('.xlsx') or fname.startswith('~$') or "PPC_NGUY" in fname.upper():
                continue
            
            fpath = os.path.join(INPUT_1_DIR, fname)
            # Identify Store
            store_name = None
            if "LHHKMT" in fname.upper(): store_name = "LHHKMT"
            elif "MUSEMORY" in fname.upper(): store_name = "MUSEMORY"
            
            if not store_name:
                continue
            
            print(f" -> Loading Bulk: {fname} (Store: {store_name})")
            date_prefix = get_date_prefix(fname)
            
            # Use pandas for fast reading and grouping
            df_bulk = pd.read_excel(fpath, sheet_name="Sponsored Products Campaigns", dtype=str)
            df_bulk.columns = [c.strip() for c in df_bulk.columns]
            
            # Convert metrics to numeric
            for col in ['Impressions', 'Clicks', 'Orders']:
                if col in df_bulk.columns:
                    df_bulk[col] = pd.to_numeric(df_bulk[col], errors='coerce').fillna(0)
                else:
                    df_bulk[col] = 0
            
            # Group by Campaign Name (Primary Key)
            # Extra: We keep Portfolio ID in grouping to allow Step 3 filtering if needed
            p_id_col = 'Portfolio ID' if 'Portfolio ID' in df_bulk.columns else ('Portfolio Id' if 'Portfolio Id' in df_bulk.columns else None)
            
            group_cols = ['Campaign Name']
            if p_id_col: group_cols.append(p_id_col)
            
            df_agg = df_bulk.groupby(group_cols, as_index=False).agg({
                'Impressions': 'sum',
                'Clicks': 'sum',
                'Orders': 'sum'
            })
            
            # Store in dict
            if store_name not in store_data_frames:
                store_data_frames[store_name] = {'df': df_agg, 'date': date_prefix, 'p_id_col': p_id_col}
            else:
                # If multiple files for same store, concatenate them
                store_data_frames[store_name]['df'] = pd.concat([store_data_frames[store_name]['df'], df_agg], ignore_index=True)

        print(f"--- Stores loaded: {list(store_data_frames.keys())}")
    except Exception as e:
        print(f"--- Error in Step 2: {e}")
        return

    # STEP 3: MAIN LOOP & TARGETED MAPPING
    print("\n[STEP 3] Running Main Loop and Mapping Data...")
    try:
        wb = openpyxl.load_workbook(master_path)
    except Exception as e:
        print(f"--- Error loading Workbook: {e}")
        return

    for target in target_list:
        sku = target['SKU']
        sku_store = target['Store']
        sku_p_id = target['Portfolio ID']
        
        print(f" -> Processing SKU: {sku} (Store: {sku_store})...", end=" ", flush=True)
        
        try:
            if sku not in wb.sheetnames:
                print("--- Warning: Sheet not found.")
                continue
                
            if sku_store not in store_data_frames:
                print(f"--- Warning: No bulk data for Store {sku_store}.")
                continue
            
            store_info = store_data_frames[sku_store]
            df_store = store_info['df']
            date_prefix = store_info['date']
            p_id_col = store_info['p_id_col']
            
            # Filter by Portfolio ID for this SKU
            if p_id_col:
                df_to_merge = df_store[df_store[p_id_col].astype(str).str.strip() == sku_p_id].copy()
            else:
                df_to_merge = df_store.copy()
            
            if df_to_merge.empty:
                print(f"--- Warning: Portfolio {sku_p_id} not found in {sku_store} data.", end=" ")
            
            # Create lookup dictionary for mapping
            lookup = {row['Campaign Name']: (row['Impressions'], row['Clicks'], row['Orders']) 
                      for _, row in df_to_merge.iterrows()}
            
            # Start openpyxl injection
            ws = wb[sku]
            headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
            
            try:
                camp_col_idx = headers.index('Campaign Name') + 1
            except ValueError:
                print("--- Error: Missing 'Campaign Name' column.")
                continue

            # Define new columns
            col_names = [f"{date_prefix}_Impression", f"{date_prefix}_Click", f"{date_prefix}_Order"]
            col_indices = []
            for c_name in col_names:
                if c_name in headers:
                    col_indices.append(headers.index(c_name) + 1)
                else:
                    new_idx = len(headers) + 1
                    ws.cell(row=1, column=new_idx, value=c_name)
                    headers.append(c_name)
                    col_indices.append(new_idx)
            
            # Update rows
            for r in range(2, ws.max_row + 1):
                camp_name = str(ws.cell(row=r, column=camp_col_idx).value).strip() if ws.cell(row=r, column=camp_col_idx).value else None
                if not camp_name: continue
                
                stats = lookup.get(camp_name, (0, 0, 0))
                ws.cell(row=r, column=col_indices[0], value=stats[0])
                ws.cell(row=r, column=col_indices[1], value=stats[1])
                ws.cell(row=r, column=col_indices[2], value=stats[2])
            
            print("Done")
            
        except Exception as e:
            print(f"--- Error processing SKU {sku}: {e}")
            continue

    # STEP 5: SAVE
    print(f"\n[STEP 5] Saving Updated Master File to {output_path}...")
    try:
        wb.save(output_path)
        print("COMPLETED SUCCESSFULY!")
    except Exception as e:
        print(f"--- Error saving file: {e}")

if __name__ == "__main__":
    main()
