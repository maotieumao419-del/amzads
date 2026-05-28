# main.py
"""
Pipeline chính cho UI Dashboard.

Luồng:
  1. Tìm file PPC_*_UPDATED.xlsx mới nhất từ D_auto_bulk_updater/data/final_xlsx/
  2. Đọc và parse toàn bộ dữ liệu (keyword-level + date blocks)
  3. Tổng hợp lên campaign-level
  4. Build Dashboard Excel đa sheet
  5. Xuất ra UI/data/output/AMZ_Interactive_Dashboard.xlsx
"""

import os
import sys
import pandas as pd
from datetime import datetime

# Thêm thư mục UI vào path để import được các module
UI_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, UI_DIR)

from ui_data_prep import (
    find_latest_updated_file,
    load_ppc_updated,
    aggregate_campaigns,
    aggregate_campaign_ts,
)
from ui_builder import build_interactive_dashboard
from config import INPUT_SUBPATH, INPUT_PATTERN, OUTPUT_SUBPATH, OUTPUT_FILE


def main():
    print("=" * 60)
    print("  AMAZON ADS DASHBOARD GENERATOR")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # ── 1. Xác định đường dẫn ─────────────────────────────────────────────
    BASE_DIR = os.path.dirname(UI_DIR)   # TEST/
    data_dir = os.path.join(BASE_DIR, "D_auto_bulk_updater", INPUT_SUBPATH)

    if not os.path.exists(data_dir):
        # Fallback: tìm trong thư mục hiện tại
        data_dir = os.path.join(UI_DIR, "data", "input")
        os.makedirs(data_dir, exist_ok=True)
        print(f"  Th mc D_auto_bulk_updater khng tm thy.")
        print(f"   ang tm trong: {data_dir}")

    # ── 2. Tìm file mới nhất ──────────────────────────────────────────────
    file_path = find_latest_updated_file(data_dir, INPUT_PATTERN)

    if not file_path:
        print(f"\n Khng tm thy file PPC_*_UPDATED.xlsx no trong:")
        print(f"   {data_dir}")
        print("\nHng dn:")
        print("  1. Chy D_auto_bulk_updater/main_pipeline.py trc")
        print("  2. Hoc copy file PPC_*_UPDATED.xlsx vo th mc trn")
        return

    file_name = os.path.basename(file_path)
    print(f"\n[FILE] File nguon: {file_name}")

    # ── 3. ETL: Đọc và xử lý dữ liệu ─────────────────────────────────────
    print("\n ang c v x l d liu...")
    etl = load_ppc_updated(file_path)

    df_kw    = etl['df_kw']
    df_kw_ts = etl['df_kw_ts']
    date_blocks = etl['date_blocks']
    skus     = etl['skus']

    if df_kw.empty:
        print(" Khng c d liu hp l t file. Kim tra li file input.")
        return

    print(f"   OK: {len(df_kw)} keyword/target tu {len(skus)} SKU")
    print(f"   Ky ngay: {', '.join(date_blocks) if date_blocks else 'Khong xac dinh'}")

    # ── 4. Tổng hợp Campaign-level ────────────────────────────────────────
    print("\n[CAMP] Dang tong hop du lieu Campaign...")
    df_camp = aggregate_campaigns(df_kw)
    print(f"   OK: {len(df_camp)} campaigns")

    # ── 5. Time-series Campaign cho Sparklines ────────────────────────────
    df_camp_ts = None
    if date_blocks and not df_kw_ts.empty:
        print(f"\n[TS] Dang tao Time-Series data ({len(date_blocks)} ky)...")
        df_camp_ts = aggregate_campaign_ts(df_kw_ts, date_blocks)
        if not df_camp_ts.empty:
            print(f"   OK: {len(df_camp_ts)} campaigns x {len(date_blocks)} ky")

    # ── 6. Build Dashboard ────────────────────────────────────────────────
    out_dir = os.path.join(UI_DIR, OUTPUT_SUBPATH)
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, OUTPUT_FILE)

    print(f"\n[BUILD] Dang tao Dashboard Excel...")

    writer   = pd.ExcelWriter(out_file, engine='xlsxwriter')
    workbook = writer.book

    etl_result = {
        'df_kw':       df_kw,
        'df_kw_ts':    df_kw_ts,
        'df_camp':     df_camp,
        'df_camp_ts':  df_camp_ts,
        'date_blocks': date_blocks,
        'skus':        skus,
        'file_name':   file_name,
    }

    build_interactive_dashboard(workbook, writer, etl_result)

    writer.close()

    # ── 7. Done ───────────────────────────────────────────────────────────
    size_kb = os.path.getsize(out_file) // 1024
    print(f"\n{'=' * 60}")
    print(f"HOAN TAT!")
    print(f"   File: {out_file}")
    print(f"   Kich thuoc: {size_kb} KB")
    print(f"   Sheets: OVERVIEW | CAMPAIGNS | KEYWORDS | DEEP DIVE")
    print(f"{'=' * 60}\n")


if __name__ == "__main__":
    main()
