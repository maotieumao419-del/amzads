import pandas as pd
import openpyxl
import sys

# Set encoding for stdout
sys.stdout.reconfigure(encoding='utf-8')

output_file = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 4\PPC_Musemory_UPDATED.xlsx'

wb = openpyxl.load_workbook(output_file, data_only=True)
ws = wb['ONM_NURSE']

# Print first few rows of columns 1-6
for r in range(1, 15):
    vals = [ws.cell(row=r, column=c).value for c in range(1, 8)]
    print(f"Row {r}: {vals}")
