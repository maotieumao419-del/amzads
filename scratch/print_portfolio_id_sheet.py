import pandas as pd
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

file_path = r"c:\Users\nnh16\ads-trading-system\TEST\Store_Musemory\Create_Campaign\data\input 1\PPC_Musemory.xlsx"
df = pd.read_excel(file_path, sheet_name="Portfolio ID")
print(df.to_string())
