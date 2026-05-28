import pandas as pd

file_path = r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\data\input\đổi tên nhé e-20260519-20260522-1779419110413.xlsx"
xl = pd.ExcelFile(file_path)

all_names = []
for sheet in xl.sheet_names:
    df = xl.parse(sheet)
    if 'Campaign Name' in df.columns:
        if 'Entity' in df.columns:
            camp_df = df[df['Entity'] == 'Campaign']
            all_names.extend(camp_df['Campaign Name'].dropna().astype(str).tolist())
        else:
            all_names.extend(df['Campaign Name'].dropna().astype(str).tolist())

unique_names = sorted(list(set(all_names)))
print(f"Total unique campaign names: {len(unique_names)}")
print("First 50 names:")
for name in unique_names[:50]:
    print(name)
