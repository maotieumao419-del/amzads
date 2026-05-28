import pandas as pd

file_path = r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\data\input\đổi tên nhé e-20260519-20260522-1779419110413.xlsx"
xl = pd.ExcelFile(file_path)

for sheet in xl.sheet_names:
    df = xl.parse(sheet)
    if 'Campaign Name' in df.columns:
        print(f"\nSheet: {sheet}")
        if 'Entity' in df.columns:
            camp_df = df[df['Entity'] == 'Campaign']
            names = camp_df['Campaign Name'].dropna().astype(str).str.strip()
            
            renamed = names[names.str.lower().str.endswith('nguyen')]
            not_renamed = names[~names.str.lower().str.endswith('nguyen')]
            
            print(f"Total campaigns: {len(names)}")
            print(f"Successfully renamed (ends with nguyen): {len(renamed)}")
            print(f"Not renamed: {len(not_renamed)}")
            if len(renamed) > 0:
                print("Examples of renamed:")
                print(renamed.head(2).tolist())
            if len(not_renamed) > 0:
                print("Examples of NOT renamed:")
                print(not_renamed.head(2).tolist())
