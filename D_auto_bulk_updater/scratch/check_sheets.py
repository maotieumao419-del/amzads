import pandas as pd
import os

input_file = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 3\PPC_Musemory.xlsx'
xl = pd.ExcelFile(input_file)
print(f"Sheets in Input 3: {xl.sheet_names}")
