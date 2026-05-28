import pandas as pd
file_path = r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\data\input\đổi tên nhé e-20260519-20260522-1779419110413.xlsx"
out_path = r"c:\Users\nnh16\ads-trading-system\TEST\scratch\all_campaign_names.txt"

xl = pd.ExcelFile(file_path)
names = set()
for sheet in xl.sheet_names:
    df = xl.parse(sheet)
    if 'Campaign Name' in df.columns:
        if 'Entity' in df.columns:
            camp_df = df[df['Entity'] == 'Campaign']
            names.update(camp_df['Campaign Name'].dropna().astype(str).tolist())
        else:
            names.update(df['Campaign Name'].dropna().astype(str).tolist())

with open(out_path, "w", encoding="utf-8") as f:
    for n in sorted(names):
        f.write(n + "\n")
