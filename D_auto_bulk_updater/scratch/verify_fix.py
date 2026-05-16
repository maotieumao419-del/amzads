import pandas as pd
import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

output_file = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 4\PPC_Musemory_UPDATED.xlsx'
wb = openpyxl.load_workbook(output_file, data_only=True)
ws = wb['ONM_NURSE']

# Check rows 3 and 7 (which were our examples)
for r in [3, 7]:
    target = ws.cell(row=r, column=4).value
    note = ws.cell(row=r, column=5).value
    # Latest block starts at Col 7 (index 6)
    next_col = 7
    while ws.cell(row=1, column=next_col).value is not None:
        current_block = next_col
        next_col += 12
    
    metrics = [ws.cell(row=r, column=current_block+i).value for i in range(11)]
    print(f"Row {r}: Target='{target}', Note='{note}', Impressions={metrics[0]}")

# Also check bile jar-02 row 3
if 'bile jar-02' in wb.sheetnames:
    ws2 = wb['bile jar-02']
    r2 = 3
    target2 = ws2.cell(row=r2, column=4).value
    metrics2 = [ws2.cell(row=r2, column=current_block+i).value for i in range(11)]
    print(f"Sheet 'bile jar-02' Row 3: Target='{target2}', Impressions={metrics2[0]}")
