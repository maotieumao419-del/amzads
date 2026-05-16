import pandas as pd
import os
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_internal_data(internal_path):
    """
    Reads all SKU sheets from the internal file and merges them into one DataFrame.
    """
    if not os.path.exists(internal_path):
        logging.error(f"File not found: {internal_path}")
        return None

    logging.info(f"Reading internal file: {internal_path}")
    xl = pd.ExcelFile(internal_path, engine='openpyxl')
    
    all_data = []
    # Skip sheets that are not SKUs
    exclude_sheets = ['Listing', 'Portfolio ID', 'Sheet1', 'Sheet2'] 
    
    for sheet_name in xl.sheet_names:
        if sheet_name in exclude_sheets:
            continue
            
        logging.info(f"Processing SKU sheet: {sheet_name}")
        # Read starting from row 3 (header is usually on row 2 or 3 in the old file)
        # Based on previous analysis, rows 1-2 might be instructions or empty.
        # Let's read the whole sheet and clean it up.
        df = pd.read_excel(internal_path, sheet_name=sheet_name, engine='openpyxl')
        
        # Add SKU column
        df['SKU'] = sheet_name
        
        # We need to standardize column names.
        # Based on analysis: ['STT', 'Campaign Name', 'Loại Campaign', 'Target', 'Ghi chú', 'Trạng thái']
        # Plus metric columns starting from Column G.
        
        all_data.append(df)
    
    if not all_data:
        logging.warning("No SKU data found in internal file.")
        return None
        
    df_merged = pd.concat(all_data, ignore_index=True)
    
    # Ensure SKU is the first column
    cols = ['SKU'] + [c for c in df_merged.columns if c != 'SKU']
    df_merged = df_merged[cols]
    
    # Simple cleaning: remove rows where Campaign Name is NaN
    if 'Campaign Name' in df_merged.columns:
        df_merged = df_merged.dropna(subset=['Campaign Name'])
        
    return df_merged

def get_date_ranges(df):
    """
    Extracts column headers that indicate date ranges (e.g., 'Ngày 17-20').
    In the internal file, these are merged cells in the first row.
    """
    # Look for columns that start with 'Ngày' or metrics that imply a date block
    # Actually, the merged data will have column names like 'Ngày 17-20.1', 'Ngày 17-20.2' if not handled.
    # We just need the unique base names.
    import re
    date_cols = [c for c in df.columns if 'Ngày' in str(c)]
    
    unique_ranges = []
    for c in date_cols:
        # Extract the range part, e.g., '17-20' from 'Ngày 17-20'
        match = re.search(r'Ngày\s*([\d\-\/]+)', str(c))
        if match:
            date_range = match.group(1)
            if date_range not in unique_ranges:
                unique_ranges.append(date_range)
    
    return unique_ranges

if __name__ == "__main__":
    # Test run
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    INTERNAL_FILE = os.path.join(BASE_DIR, "RULE&TEMPLATE", "PPC_NGUYÊN.xlsx")
    df = get_internal_data(INTERNAL_FILE)
    if df is not None:
        print(f"Merged {len(df)} rows from all SKU sheets.")
        print("Columns found:", df.columns.tolist())
