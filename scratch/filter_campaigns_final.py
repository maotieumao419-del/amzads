# -*- coding: utf-8 -*-
import pandas as pd
import traceback
import sys
import os
import glob

# Use glob to avoid hardcoding Vietnamese characters which causes encoding issues in Windows Python
input_files = glob.glob(r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\data\input\*.xlsx")
input_file = next(f for f in input_files if '1779419110413' in f)
output_file = r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\data\input\chua_doi_ten.xlsx"

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
except Exception as e:
    with open(r"c:\Users\nnh16\ads-trading-system\TEST\scratch\error2.log", "w") as f:
        f.write(traceback.format_exc())
