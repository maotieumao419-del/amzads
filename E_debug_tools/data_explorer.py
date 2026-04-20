import pandas as pd
import os

# -------------------------------------------------------
# DEBUG: Khám phá cấu trúc PPC_NGUYÊN.xlsx & AMAQUANG
# Output: explorer_out.txt (cùng thư mục)
# -------------------------------------------------------

SCRIPT_DIR      = os.path.dirname(os.path.abspath(__file__))
RULE_DIR        = os.path.normpath(os.path.join(SCRIPT_DIR, '..', 'RULE&TEMPLATE'))
PPC_FILE        = os.path.join(RULE_DIR, 'PPC_NGUYÊN.xlsx')
AMAQUANG_FILE   = os.path.join(RULE_DIR, 'AMAQUANG+ AMATRA.xlsx')
OUTPUT_FILE     = os.path.join(SCRIPT_DIR, "explorer_out.txt")

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    try:
        f.write(f"Reading {PPC_FILE}...\n")
        ppc_xls = pd.ExcelFile(PPC_FILE)
        f.write(f"Sheets in PPC_NGUYÊN: {ppc_xls.sheet_names}\n")

        for sheet in ppc_xls.sheet_names[2:5]:
            df = pd.read_excel(ppc_xls, sheet_name=sheet)
            f.write(f"\n--- Sheet: {sheet} ---\n")
            f.write(f"Columns: {df.columns.tolist()}\n")
            f.write("First 2 targets:\n")
            if 'Target' in df.columns:
                f.write(df['Target'].head(2).to_string() + "\n")
            else:
                f.write("No 'Target' column found\n")
            f.write(f"Trạng thái index: {df.columns.tolist().index('Trạng thái') if 'Trạng thái' in df.columns else 'Not found'}\n")

        f.write(f"\nReading {AMAQUANG_FILE}...\n")
        ama_xls = pd.ExcelFile(AMAQUANG_FILE)
        f.write(f"Sheets in AMAQUANG: {ama_xls.sheet_names}\n")

        sheet_nm = 'Sponsored Products Campaigns'
        if sheet_nm in ama_xls.sheet_names:
            df_ama = pd.read_excel(ama_xls, sheet_name=sheet_nm)
            f.write(f"Columns in AMAQUANG: {df_ama.columns.tolist()}\n")
            f.write("First 10 campaigns:\n")
            camp_df = df_ama[df_ama['Entity'] == 'Campaign']
            f.write(camp_df[['Campaign Name', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders']].head(10).to_string() + "\n")

    except Exception as e:
        f.write(f"Error: {e}\n")

print(f"Done. Output: {OUTPUT_FILE}")
