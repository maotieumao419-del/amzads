import pandas as pd
import os

# -------------------------------------------------------
# NHIỆM VỤ: Bước 1 của pipeline A_create_campaign
# Đọc file Excel từ data/input/report.xlsx
# Xuất thành CSV raw: data/input/raw_sp_date.csv
# -------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def main():
    input_file  = os.path.join(SCRIPT_DIR, "data", "input", "report.xlsx")
    output_file = os.path.join(SCRIPT_DIR, "data", "input", "raw_sp_date.csv")

    print(f"Reading {input_file}...")
    try:
        df = pd.read_excel(
            input_file,
            sheet_name='Sponsored Products Campaigns',
            dtype=str,
            engine='openpyxl'
        )
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        df.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"Success! Extracted {len(df)} rows and {len(df.columns)} columns.")
        print(f"Saved raw data to {output_file}")
    except Exception as e:
        print(f"Error processing the file: {e}")

if __name__ == "__main__":
    main()
