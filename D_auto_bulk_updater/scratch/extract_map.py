import pandas as pd
import json

file_path = r'C:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\raw_xlsx\PPC_Musemory.xlsx'
df = pd.read_excel(file_path, sheet_name='Listing', engine='openpyxl')
mapping = dict(zip(df['SKU'].astype(str), df['ASIN'].astype(str)))
print(json.dumps(mapping, indent=2))
