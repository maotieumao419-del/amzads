import pandas as pd
df = pd.read_excel(r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\C_mass_sop_factory\data\output\Bulk_Create_250TH_DEVOM_DAIBANG1776-2026.xlsx")
with open("test_output.txt", "w", encoding="utf-8") as f:
    f.write(df[['Entity', 'Campaign Name', 'Ad Group Name', 'Keyword Text']].head(15).to_string())
    f.write(f"\nTotal rows: {len(df)}")
