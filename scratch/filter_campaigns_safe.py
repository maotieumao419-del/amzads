import pandas as pd
import traceback
import sys
import os

input_file = r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\data\input\đổi tên nhé e-20260519-20260522-1779419110413.xlsx"
output_file = r"C:\Users\nnh16\.gemini\antigravity-ide\brain\d45ed5a2-5819-4821-8151-c795028fb456\scratch\chua_doi_ten.xlsx"
log_file = r"C:\Users\nnh16\.gemini\antigravity-ide\brain\d45ed5a2-5819-4821-8151-c795028fb456\scratch\error.log"

os.makedirs(os.path.dirname(output_file), exist_ok=True)

try:
    xl = pd.ExcelFile(input_file)
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            if 'Campaign Name' in df.columns:
                renamed_cids = set()
                if 'Entity' in df.columns:
                    camp_df = df[df['Entity'].astype(str).str.strip().str.lower() == 'campaign']
                    for idx, row in camp_df.iterrows():
                        cname = str(row['Campaign Name']).lower()
                        if 'nguyen' in cname or '_sp0' in cname:
                            cid = row.get('Campaign ID')
                            if pd.notna(cid):
                                renamed_cids.add(cid)
                for idx, row in df.iterrows():
                    cname = str(row['Campaign Name']).lower()
                    if 'nguyen' in cname or '_sp0' in cname:
                        cid = row.get('Campaign ID')
                        if pd.notna(cid):
                            renamed_cids.add(cid)
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
            else:
                df.to_excel(writer, sheet_name=sheet, index=False)
    with open(log_file, "w") as f:
        f.write("Success")
except Exception as e:
    with open(log_file, "w") as f:
        f.write(traceback.format_exc())
