import os
import sys
import json
import pandas as pd

# Thêm thư mục UI vào path để import các module
UI_DIR = os.path.dirname(os.path.abspath(__file__))
if UI_DIR not in sys.path:
    sys.path.insert(0, UI_DIR)

from ui_data_prep import (
    find_latest_updated_file,
    load_ppc_updated,
    aggregate_campaigns,
    aggregate_campaign_ts,
)
from config import INPUT_SUBPATH, INPUT_PATTERN

def export_dashboard_data():
    print("=" * 60)
    print("  AMAZON ADS - STATIC JSON EXPORTER")
    print("=" * 60)

    BASE_DIR = os.path.dirname(UI_DIR)   # TEST/
    data_dir = os.path.join(BASE_DIR, "D_auto_bulk_updater", INPUT_SUBPATH)

    if not os.path.exists(data_dir):
        data_dir = os.path.join(UI_DIR, "data", "input")

    file_path = find_latest_updated_file(data_dir, INPUT_PATTERN)

    if not file_path:
        print("ERROR: Cannot find source file (PPC_*_UPDATED.xlsx)")
        sys.exit(1)

    print(f"Processing file: {os.path.basename(file_path)}...")

    try:
        # Tái sử dụng ETL
        etl = load_ppc_updated(file_path)
        df_kw = etl['df_kw']
        df_kw_ts = etl['df_kw_ts']
        date_blocks = etl['date_blocks']
        skus = etl['skus']

        if df_kw.empty:
            print("ERROR: File has no valid data")
            sys.exit(1)

        # Tổng hợp Campaign
        df_camp = aggregate_campaigns(df_kw)

        # Tổng hợp Time Series cho Campaign
        df_camp_ts = pd.DataFrame()
        if date_blocks and not df_kw_ts.empty:
            df_camp_ts = aggregate_campaign_ts(df_kw_ts, date_blocks)

        # Chuyển đổi DataFrame sang JSON serializable (Điền NaN thành 0)
        df_kw = df_kw.fillna(0)
        df_camp = df_camp.fillna(0)
        df_camp_ts = df_camp_ts.fillna(0)

        # Cấu trúc JSON
        data = {
            "source_file": os.path.basename(file_path),
            "date_blocks": date_blocks,
            "skus": skus,
            "summary": {
                "total_spend": float(df_camp['Spend'].sum()),
                "total_sales": float(df_camp['Sales'].sum()),
                "total_orders": int(df_camp['Orders'].sum()),
                "total_clicks": int(df_camp['Clicks'].sum()),
                "total_impressions": int(df_camp['Impressions'].sum()),
                "overall_acos": float(df_camp['Spend'].sum() / df_camp['Sales'].sum()) if df_camp['Sales'].sum() > 0 else 0.0,
            },
            "campaigns": df_camp.to_dict(orient="records"),
            "campaign_time_series": df_camp_ts.reset_index().to_dict(orient="records") if not df_camp_ts.empty else [],
            "keywords": df_kw.to_dict(orient="records")
        }

        # Ghi file ra thư mục webapp
        out_path = os.path.join(UI_DIR, "webapp", "src", "data.json")
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        print(f"\n[SUCCESS] Data exported to local JSON file:")
        print(f"-> {out_path}")
        print("Web App is now updated. You can deploy it to Vercel!")

    except Exception as e:
        print(f"SYSTEM ERROR: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    export_dashboard_data()
