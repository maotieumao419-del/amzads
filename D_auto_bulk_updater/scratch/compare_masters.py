import pandas as pd
import os

path1 = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 1\BulkSheetExport_3004-0405.xlsx'
path2 = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\input 2\BulkSheetExport_3004-0405.xlsx'

df1 = pd.read_excel(path1, sheet_name='Sponsored Products Campaigns', engine='openpyxl')
df2 = pd.read_excel(path2, sheet_name='Sponsored Products Campaigns', engine='openpyxl')

print(f"Input 1: {len(df1)} rows")
print(f"Input 2: {len(df2)} rows")

# Check if Campaign 'B0F62HNF16_ ONM_NURSE_KT_nurse' exists in both
print(f"\nCampaign 'B0F62HNF16_ ONM_NURSE_KT_nurse' in Input 1: {len(df1[df1['Campaign Name'].astype(str).str.contains('ONM_NURSE', na=False)])} rows")
print(f"Campaign 'B0F62HNF16_ ONM_NURSE_KT_nurse' in Input 2: {len(df2[df2['Campaign Name'].astype(str).str.contains('ONM_NURSE', na=False)])} rows")
