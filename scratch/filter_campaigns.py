import pandas as pd
import os

input_file = r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\data\input\đổi tên nhé e-20260519-20260522-1779419110413.xlsx"
output_file = r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\data\input\chua_doi_ten.xlsx"

xl = pd.ExcelFile(input_file)
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        if 'Campaign Name' in df.columns:
            # Find Campaign IDs of successfully renamed campaigns
            # We consider it renamed if it contains 'nguyen' or '_SP0'
            renamed_cids = set()
            
            # Look at Entity == Campaign first
            if 'Entity' in df.columns:
                camp_df = df[df['Entity'].astype(str).str.strip().str.lower() == 'campaign']
                for idx, row in camp_df.iterrows():
                    cname = str(row['Campaign Name']).lower()
                    if 'nguyen' in cname or '_sp0' in cname:
                        cid = row.get('Campaign ID')
                        if pd.notna(cid):
                            renamed_cids.add(cid)
            
            # Also check any row just in case
            for idx, row in df.iterrows():
                cname = str(row['Campaign Name']).lower()
                if 'nguyen' in cname or '_sp0' in cname:
                    cid = row.get('Campaign ID')
                    if pd.notna(cid):
                        renamed_cids.add(cid)
            
            # Now filter the dataframe
            # Keep rows where Campaign ID is not in renamed_cids
            # Or if Campaign ID is null, check Campaign Name directly
            def is_not_renamed(row):
                cid = row.get('Campaign ID')
                if pd.notna(cid) and cid in renamed_cids:
                    return False
                cname = str(row.get('Campaign Name', '')).lower()
                if 'nguyen' in cname or '_sp0' in cname:
                    return False
                return True
                
            filtered_df = df[df.apply(is_not_renamed, axis=1)]
            filtered_df.to_excel(writer, sheet_name=sheet, index=False)
            print(f"Sheet {sheet}: kept {len(filtered_df)} out of {len(df)} rows.")
        else:
            df.to_excel(writer, sheet_name=sheet, index=False)
            print(f"Sheet {sheet}: no Campaign Name column, kept {len(df)} rows.")

print(f"Created {output_file}")
