import os
import glob
import json
import pandas as pd
from data_prep import process_data
from sheet_overview import build_overview_sheet
from sheet_campaigns import build_campaigns_sheet
from sheet_keywords import build_keywords_sheet
from sheet_search_terms import build_search_terms_sheet
from data_time_series import process_time_series
from charts import add_sparkline_sheet, draw_campaign_sparklines
from sheet_deep_dive import build_deep_dive_sheet

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
CONFIG_DIR = os.path.join(BASE_DIR, "config")

def load_configs():
    with open(os.path.join(CONFIG_DIR, "Rule_Engine.json"), "r", encoding="utf-8") as f:
        rules = json.load(f)
    with open(os.path.join(CONFIG_DIR, "Season_Calendar.json"), "r", encoding="utf-8") as f:
        calendar = json.load(f)
    return rules, calendar

def main():
    print("Bat dau tao Dashboard...")
    rules, calendar = load_configs()
    
    input_files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
    if not input_files:
        print("Khong tim thay file .xlsx nao trong thu muc input/")
        return
    
    file_path = input_files[0]
    print(f"Dang doc file: {os.path.basename(file_path)}")
    
    # 1. Process Data
    df_camp, df_kw, df_st, today = process_data(file_path, calendar)
    df_ts, df_ts_agg = process_time_series(df_camp, days=14)

    # 2. Setup Excel Writer
    out_file = os.path.join(OUTPUT_DIR, f"Amazon_Ads_Dashboard_{today.strftime('%Y%m%d_%H%M%S')}.xlsx")
    writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
    workbook = writer.book

    # 3. Build Sheets
    df_ts_agg_saved = add_sparkline_sheet(workbook, writer, df_ts_agg)
    build_overview_sheet(workbook, writer, df_camp, df_kw, today)
    df_camp_out = build_campaigns_sheet(writer, df_camp)
    build_keywords_sheet(writer, df_kw)
    build_search_terms_sheet(writer, df_st)
    
    build_deep_dive_sheet(workbook, writer, df_ts, df_camp_out)
    
    # Kích hoạt sheet OVERVIEW để khi mở file sẽ thấy nó đầu tiên
    writer.sheets['📊 OVERVIEW'].activate()
    
    # 4. Draw Charts
    draw_campaign_sparklines(workbook, writer, df_camp_out, df_ts_agg_saved)

    # 5. Save
    writer.close()
    print(f"Da tao thanh cong Dashboard: {out_file}")

if __name__ == "__main__":
    main()
