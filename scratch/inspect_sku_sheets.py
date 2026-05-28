import pandas as pd
import openpyxl
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

file_path = r"c:\Users\nnh16\ads-trading-system\TEST\Store_Musemory\Create_Campaign\data\input 1\PPC_Musemory.xlsx"
print("Loading workbook...")
wb = openpyxl.load_workbook(file_path, read_only=True)
sheets = [s for s in wb.sheetnames if s not in ("Listing", "Portfolio ID")]

print("SKU Sheets in PPC_Musemory.xlsx:", sheets)

for s in sheets[:3]:
    print(f"\nSheet {s}:")
    df = pd.read_excel(file_path, sheet_name=s)
    print("Columns:", [str(c) for c in df.columns])
    print(df.head(3))
