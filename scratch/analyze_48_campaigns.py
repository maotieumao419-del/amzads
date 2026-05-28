import os
import re
import pandas as pd

def main():
    test_dir = r"c:\Users\nnh16\ads-trading-system\TEST"
    bulk_path = os.path.join(test_dir, "bulk-a2916r2qstxqjc-20260518-20260521-1779350870446.xlsx")
    master_path = os.path.join(test_dir, "D_auto_bulk_updater", "data", "output", "amazon_ready", "MASTER_RENAME_BULK_19052026.xlsx")
    unrenamed_path = os.path.join(test_dir, "G_reset_system", "unrenamed_campaigns.xlsx")

    # 1. Load data
    print("Loading unrenamed campaigns...")
    df_unrenamed = pd.read_excel(unrenamed_path, sheet_name="Sponsored Products Campaigns", dtype=str)
    
    # 2. Get unique unrenamed campaign IDs
    unrenamed_cids = set()
    campaign_name_by_id = {}
    for idx, row in df_unrenamed.iterrows():
        entity = str(row.get("Entity", "")).strip().lower()
        cid = str(row.get("Campaign ID", "")).strip()
        cname = str(row.get("Campaign Name (Informational only)", row.get("Campaign Name", ""))).strip()
        if cid and cid.lower() != 'nan':
            unrenamed_cids.add(cid)
            if entity == 'campaign':
                campaign_name_by_id[cid] = cname
                
    print(f"Loaded {len(unrenamed_cids)} unique campaign IDs from unrenamed_campaigns.xlsx.")
    
    # 3. Load MASTER_RENAME_BULK_19052026.xlsx to see which ones were in the rename file
    master_rename_cids = set()
    master_rename_names = {}
    if os.path.exists(master_path):
        df_master = pd.read_excel(master_path, dtype=str)
        df_master.columns = [str(c).strip() for c in df_master.columns]
        for idx, row in df_master.iterrows():
            cid = str(row.get("Campaign ID", "")).strip()
            entity = str(row.get("Entity", "")).strip().lower()
            cname = str(row.get("Campaign Name", "")).strip()
            if cid and cid.lower() != 'nan':
                master_rename_cids.add(cid)
                if entity == 'campaign':
                    master_rename_names[cid] = cname
        print(f"Loaded {len(master_rename_cids)} campaigns from MASTER_RENAME_BULK_19052026.xlsx.")
    else:
        print("MASTER_RENAME_BULK_19052026.xlsx not found!")

    # 4. Separate unrenamed campaigns into:
    #    - Group 1: Campaigns that WERE in MASTER_RENAME_BULK_19052026.xlsx (but failed to update or have special names)
    #    - Group 2: Campaigns that WERE NOT in MASTER_RENAME_BULK_19052026.xlsx (skipped by engine)
    in_master = unrenamed_cids.intersection(master_rename_cids)
    not_in_master = unrenamed_cids - master_rename_cids

    print(f"\nGroup 1: In Master Rename File but still unrenamed ({len(in_master)} campaigns):")
    for cid in in_master:
        print(f"  ID: {cid} | Current: '{campaign_name_by_id.get(cid, 'Unknown')}' | Master Rename Proposed: '{master_rename_names.get(cid, 'Unknown')}'")

    print(f"\nGroup 2: Not in Master Rename File ({len(not_in_master)} campaigns):")
    
    # Let's inspect Group 2 details in unrenamed campaigns to find why they were skipped
    # Group rows by campaign ID for detailed analysis
    grouped = df_unrenamed.groupby("Campaign ID")
    
    for cid in not_in_master:
        if cid not in campaign_name_by_id:
            continue
        cname = campaign_name_by_id[cid]
        crows = grouped.get_group(cid)
        entities = crows["Entity"].str.strip().str.lower().tolist()
        
        sku = ""
        asin = ""
        ad_rows = crows[crows["Entity"].str.strip().str.lower() == 'product ad']
        if not ad_rows.empty:
            sku = str(ad_rows.iloc[0].get("SKU", ""))
            asin = str(ad_rows.iloc[0].get("ASIN (Informational only)", ad_rows.iloc[0].get("ASIN", "")))
            
        kws = crows[crows["Entity"].str.strip().str.lower() == 'keyword']["Keyword Text"].dropna().tolist()
        pts = crows[crows["Entity"].str.strip().str.lower() == 'product targeting']["Product Targeting Expression"].dropna().tolist()
        
        print(f"  ID: {cid} | Name: '{cname}'")
        print(f"    Entities: {dict(pd.Series(entities).value_counts())}")
        print(f"    SKU: '{sku}' | ASIN: '{asin}'")
        if kws:
            print(f"    Keywords (up to 3): {kws[:3]}")
        if pts:
            print(f"    Product Targets (up to 3): {pts[:3]}")

if __name__ == '__main__':
    main()
