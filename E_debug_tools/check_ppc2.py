import pandas as pd
import os

# -------------------------------------------------------
# DEBUG: Khám phá cột Campaign Name trong PPC_NGUYÊN.xlsx
# Output: explorer_ppc2.txt (cùng thư mục)
# -------------------------------------------------------

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
PPC_FILE    = os.path.normpath(os.path.join(SCRIPT_DIR, '..', 'RULE&TEMPLATE', 'PPC_NGUYÊN.xlsx'))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "explorer_ppc2.txt")

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    try:
        ppc_xls = pd.ExcelFile(PPC_FILE)
        for sheet in ppc_xls.sheet_names[2:5]:
            df = pd.read_excel(ppc_xls, sheet_name=sheet)
            f.write(f"--- Sheet: {sheet} ---\n")
            if 'Campaign Name' in df.columns:
                f.write("Campaign Names:\n")
                f.write(df['Campaign Name'].head(5).to_string() + "\n")
            if 'Trạng thái' in df.columns:
                f.write(f"Columns around Trạng thái: {df.columns.tolist()[-5:]}\n")
    except Exception as e:
        f.write(f"Error: {e}\n")

print(f"Done. Output: {OUTPUT_FILE}")
