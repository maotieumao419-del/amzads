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
from sheet_daily_summary import build_daily_summary_sheet

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
    
    import sys
    DB_DIR = os.path.join(os.path.dirname(BASE_DIR), "database", "code")
    sys.path.append(DB_DIR)
    from db_manager import ingest_bulk_to_db
    DB_PATH = os.path.join(os.path.dirname(BASE_DIR), "database", "amazon_ads_history.db")
    
    PARSE_DIR = os.path.join(os.path.dirname(BASE_DIR), "phase_bo_sung", "code")
    sys.path.append(PARSE_DIR)
    from parse_filename import parse_bulk_filename
    from datetime import datetime
    
    input_files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
    
    if not input_files:
        print("Khong tim thay file .xlsx nao trong thu muc input/")
        return
        
    print(f"Tim thay {len(input_files)} file bulk. Dang xu ly va dua vao database...")
    
    file_metadata = []
    for f_path in input_files:
        meta = parse_bulk_filename(f_path)
        df_camp, df_kw, df_st, _today = process_data(f_path, calendar)
        start_date = meta.get("start_date")
        end_date = meta.get("end_date")
        days = meta.get("days_duration")
        
        if start_date and end_date and days:
            ingest_bulk_to_db(DB_PATH, df_camp, start_date, end_date, days)
            file_end_dt = datetime.strptime(end_date, "%Y-%m-%d").date()
        else:
            file_end_dt = _today.date() if isinstance(_today, datetime) else _today
            
        file_metadata.append({
            'file_path': f_path,
            'start_date_str': start_date,
            'end_date_str': end_date,
            'end_date': file_end_dt,
            'df_camp': df_camp,
            'df_kw': df_kw,
            'df_st': df_st,
            'today': file_end_dt
        })
        print(f" - Da nap: {os.path.basename(f_path)}")
        
    # Chon file moi nhat de lam Dashboard hien tai
    file_metadata.sort(key=lambda x: x['end_date'], reverse=True)
    recent_file = file_metadata[0]
    
    file_path = recent_file['file_path']
    df_camp = recent_file['df_camp']
    df_kw = recent_file['df_kw']
    df_st = recent_file['df_st']
    today = recent_file['today']
    
    dashboard_days = 21
    data_type = 'daily'
    
    print(f"\nChon file gan nhat de tao Dashboard: {os.path.basename(file_path)}")
    print(f"Lay lich su {dashboard_days} ngay gan nhat (den {today})...")
    
    df_ts, df_ts_agg = process_time_series(df_camp, today=today, days=dashboard_days)
    
    # Gom nhóm theo Date cho sheet Daily Summary
    df_daily_summary = df_ts.groupby('Date')[['Spend', 'Sales', 'Orders', 'Impressions', 'Clicks']].sum().reset_index()
    
    # Tao df_ts_bulk cho sheet Deep Dive de khong bi chia trung binh theo ngay
    df_ts_bulk_list = []
    df_kw_bulk_list = []
    for meta in file_metadata:
        df = meta['df_camp'].copy()
        df_kw_meta = meta['df_kw'].copy()
        s_date = meta.get('start_date_str')
        e_date = meta.get('end_date_str')
        if s_date and e_date and s_date != e_date:
            date_str = f"{s_date} to {e_date}"
        elif s_date:
            date_str = str(s_date)
        else:
            date_str = str(meta['today'])
        df['Date'] = date_str
        df_kw_meta['Date'] = date_str
        df_ts_bulk_list.append(df)
        df_kw_bulk_list.append(df_kw_meta)
        
    df_ts_bulk = pd.concat(df_ts_bulk_list, ignore_index=True)
    df_kw_bulk = pd.concat(df_kw_bulk_list, ignore_index=True)

    # 2. Setup Excel Writer
    out_file = os.path.join(OUTPUT_DIR, f"Amazon_Ads_Dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
    workbook = writer.book

    # 3. Build Sheets
    df_ts_agg_saved = add_sparkline_sheet(workbook, writer, df_ts_agg)
    build_overview_sheet(workbook, writer, df_camp, df_kw, today)
    df_camp_out = build_campaigns_sheet(writer, df_camp)
    build_keywords_sheet(writer, df_kw)
    build_search_terms_sheet(writer, df_st)
    build_daily_summary_sheet(workbook, writer, df_daily_summary)
    
    build_deep_dive_sheet(workbook, writer, df_ts_bulk, df_camp_out, df_kw_bulk, data_type='bulk')
    
    # Kích hoạt sheet OVERVIEW để khi mở file sẽ thấy nó đầu tiên
    writer.sheets['📊 OVERVIEW'].activate()
    
    # 4. Draw Charts
    draw_campaign_sparklines(workbook, writer, df_camp_out, df_ts_agg_saved)

    # 5. Save
    writer.close()
    print(f"Da tao thanh cong Dashboard: {out_file}")

if __name__ == "__main__":
    main()
