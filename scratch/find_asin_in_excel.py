import pandas as pd
import re

file_path = r"c:\Users\nnh16\ads-trading-system\TEST\Store_Musemory\Create_Campaign\data\input 1\PPC_Musemory.xlsx"
xls = pd.ExcelFile(file_path)

asin_pattern = re.compile(r"\bB[A-Z0-9]{9}\b", re.IGNORECASE)

for sheet_name in xls.sheet_names:
    print(f"Scanning sheet: {sheet_name}")
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    for col in df.columns:
        for idx, val in df[col].dropna().items():
            s_val = str(val)
            matches = asin_pattern.findall(s_val)
            if matches:
                print(f"  Found {matches} in column {col}, row {idx}: {repr(s_val)}")
