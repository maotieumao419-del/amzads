import pandas as pd
import os

path2 = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 2\BulkSheetExport_3004-0405.xlsx'

if os.path.exists(path2):
    df = pd.read_excel(path2, sheet_name='Sponsored Products Campaigns', engine='openpyxl', nrows=0)
    print("Columns in input 2 file:")
    print(df.columns.tolist())
else:
    print("File not found in input 2.")
