import pandas as pd
import openpyxl
import sys

# Configure stdout to handle UTF-8 printing
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

file_path = r"c:\Users\nnh16\ads-trading-system\TEST\Store_Musemory\Create_Campaign\data\input 1\PPC_Musemory.xlsx"
print("Loading workbook...")

df = pd.read_excel(file_path, sheet_name="Listing")
print("Listing sheet columns:")
print([str(c) for c in df.columns])

print("\nListing sheet first 10 rows:")
for idx, row in df.head(10).iterrows():
    print(f"Row {idx}: SKU={repr(row.get('SKU'))}, Portfolio ID={repr(row.get('Portfolio Id'))}, Status={repr(row.get('Status'))}, Store={repr(row.get('Store') if 'Store' in row else None)}")
