# ui_builder.py
"""
View Controller / Orchestrator.
Gọi tuần tự các sheet builders và chart builders để tạo Dashboard hoàn chỉnh.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import SHEET_OVERVIEW, TARGET_ACOS
from ui_charts import add_ts_data_sheet, draw_campaign_sparklines
from sheets.sheet_overview import build_overview_sheet
from sheets.sheet_campaigns import build_campaigns_sheet
from sheets.sheet_keywords import build_keywords_sheet
from sheets.sheet_deep_dive import build_deep_dive_sheet


def build_interactive_dashboard(workbook, writer, etl_result: dict):
    """
    Orchestrator chính.
    etl_result = {
      'df_kw':       DataFrame keyword-level,
      'df_kw_ts':    DataFrame keyword time-series wide format,
      'df_camp':     DataFrame campaign-level aggregated,
      'df_camp_ts':  DataFrame campaign time-series (pivot),
      'date_blocks': list[str] — tất cả kỳ ngày,
      'skus':        list[str],
      'file_name':   str — tên file nguồn,
    }
    """
    df_kw      = etl_result['df_kw']
    df_kw_ts   = etl_result['df_kw_ts']
    df_camp    = etl_result['df_camp']
    df_camp_ts = etl_result.get('df_camp_ts', None)
    date_blocks = etl_result['date_blocks']
    skus       = etl_result['skus']
    file_name  = etl_result.get('file_name', '')

    # ── 1. Sheet ẩn TS_DATA (Sparkline data) ─────────────────────────────
    df_ts_agg = None
    if df_camp_ts is not None and not df_camp_ts.empty:
        df_ts_agg = add_ts_data_sheet(workbook, writer, df_camp_ts)

    # ── 2. OVERVIEW ───────────────────────────────────────────────────────
    build_overview_sheet(workbook, writer, df_camp, df_kw, skus, file_name)

    # ── 3. CAMPAIGNS ──────────────────────────────────────────────────────
    df_camp_out, cols_out = build_campaigns_sheet(workbook, writer, df_camp)

    # ── 4. KEYWORDS ───────────────────────────────────────────────────────
    build_keywords_sheet(workbook, writer, df_kw)

    # ── 5. DEEP DIVE ──────────────────────────────────────────────────────
    build_deep_dive_sheet(workbook, writer, df_kw, df_camp, date_blocks, skus)

    # ── 6. Sparklines (sau khi tất cả sheets đã được tạo) ────────────────
    if df_ts_agg is not None:
        draw_campaign_sparklines(workbook, writer, df_camp_out, df_ts_agg)

    # ── 7. Activate sheet OVERVIEW khi mở file ───────────────────────────
    writer.sheets[SHEET_OVERVIEW].activate()
