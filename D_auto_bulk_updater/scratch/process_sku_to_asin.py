import pandas as pd
import glob
import os
import re

def extract_asin(text):
    if pd.isna(text):
        return None
    match = re.search(r'(B[A-Z0-9]{9})', str(text).upper())
    if match:
        return match.group(1)
    return None

def main():
    # 1. Load Mapping
    mapping_file = r'C:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\raw_xlsx\PPC_Musemory.xlsx'
    df_map = pd.read_excel(mapping_file, sheet_name='Listing', engine='openpyxl')
    
    sku_to_asin = {}
    for _, row in df_map.iterrows():
        sku = str(row.get('SKU', '')).strip()
        link = str(row.get('Link', '')).strip()
        
        asin = extract_asin(link)
        if not asin:
            asin = extract_asin(sku)
            
        if sku and asin and sku != 'nan' and asin != 'nan':
            sku_to_asin[sku] = asin
            
    # 2. Process Files
    input_folder = r"C:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\output\update"
    output_folder = r"C:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\output\amazon_ready"
    os.makedirs(output_folder, exist_ok=True)
    
    files = glob.glob(os.path.join(input_folder, "*.xlsx"))
    print(f"Found {len(files)} files to process.")
    
    success_count = 0
    missing_mapping = set()
    
    for f in files:
        basename = os.path.basename(f)
        if basename.startswith("~"):
            continue
            
        sku_from_file = basename.replace("_campaign_updates.xlsx", "").strip()
        
        asin = None
        for k, v in sku_to_asin.items():
            if k.lower() == sku_from_file.lower():
                asin = v
                break
                
        if not asin:
            missing_mapping.add(sku_from_file)
            print(f"WARNING: No ASIN mapping found for SKU: {sku_from_file}")
            continue
            
        try:
            df = pd.read_excel(f, engine='openpyxl')
            if 'Campaign Name' in df.columns:
                # Replace the exact SKU (case-insensitive) with ASIN
                pattern = re.compile(re.escape(sku_from_file), re.IGNORECASE)
                df['Campaign Name'] = df['Campaign Name'].apply(lambda x: pattern.sub(asin, str(x)) if pd.notnull(x) else x)
                
                # If there's Ad Group Name
                if 'Ad Group Name' in df.columns:
                    df['Ad Group Name'] = df['Ad Group Name'].apply(lambda x: pattern.sub(asin, str(x)) if pd.notnull(x) else x)

            out_file = os.path.join(output_folder, f"{asin}_campaign_updates.xlsx")
            df.to_excel(out_file, index=False, engine='openpyxl')
            success_count += 1
            print(f"Processed: {basename} -> {os.path.basename(out_file)}")
        except Exception as e:
            print(f"Error processing {basename}: {e}")
            
    print(f"Successfully processed {success_count} files.")
    if missing_mapping:
        print(f"Missing mappings for {len(missing_mapping)} SKUs.")

if __name__ == '__main__':
    main()
