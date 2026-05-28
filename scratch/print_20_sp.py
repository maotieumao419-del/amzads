import pandas as pd

file_path = r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\data\input\đổi tên nhé e-20260519-20260522-1779419110413.xlsx"
df = pd.read_excel(file_path, sheet_name='Sponsored Products Campaigns')

if 'Campaign Name' in df.columns:
    if 'Entity' in df.columns:
        names = df[df['Entity'] == 'Campaign']['Campaign Name'].dropna().astype(str).tolist()
    else:
        names = df['Campaign Name'].dropna().astype(str).tolist()
    
    unique_names = list(set(names))
    print(f"Unique campaigns in SP: {len(unique_names)}")
    for n in unique_names[:20]:
        print(n)
