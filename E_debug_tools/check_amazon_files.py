import pandas as pd
import os

# -------------------------------------------------------
# DEBUG: Khám phá cấu trúc các file Excel liên quan
# Đọc: Final_Mass_Update_Upload.xlsx, QC_Nguyên TDH.xlsx
# Output: explorer_amaz.txt (cùng thư mục)
# -------------------------------------------------------

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
RULE_DIR    = os.path.normpath(os.path.join(SCRIPT_DIR, '..', 'RULE&TEMPLATE'))
D_DIR       = os.path.normpath(os.path.join(SCRIPT_DIR, '..', 'D_auto_bulk_updater'))
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "explorer_amaz.txt")

FILES_TO_CHECK = [
    os.path.join(D_DIR,   'data', 'output', 'Final_Mass_Update_Upload.xlsx'),
    os.path.join(RULE_DIR, 'QC_Nguyên TDH.xlsx'),
]

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    for file_path in FILES_TO_CHECK:
        try:
            xls = pd.ExcelFile(file_path)
            f.write(f"File: {os.path.basename(file_path)} - Sheets: {xls.sheet_names}\n")
            if 'Sponsored Products Campaigns' in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name='Sponsored Products Campaigns')
                f.write(f"Columns: {df.columns.tolist()[:15]}...\n")
                if 'Campaign Name' in df.columns:
                    f.write(df[['Campaign Name']].head(3).to_string() + "\n")
        except Exception as e:
            f.write(f"Error {os.path.basename(file_path)}: {e}\n")

print(f"Done. Output: {OUTPUT_FILE}")
