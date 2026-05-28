# ui_data_prep.py
"""
ETL & Business Logic Layer cho UI Dashboard.
Đọc file PPC_*_UPDATED.xlsx (file mới nhất) và trả về các DataFrame đã xử lý.

Cấu trúc PPC_UPDATED.xlsx:
  - Mỗi sheet = 1 SKU
  - Row 1: merged title "SKU: <name>"
  - Row 2: headers (STT, Campaign Name, Loại Campaign, Target, Ghi chú, Status,
            [Ngày XXXX → Impressions, Clicks, CTR, Spend, Sales, Orders, Units, CVR, ACOS, CPC, ROAS, Placement Note], ...)
  - Row 3+: data rows
"""

import os
import re
import glob
import logging
import pandas as pd

from config import (
    TARGET_ACOS, HEALTH_ACOS_STAR, HEALTH_ACOS_BLEEDER,
    HEALTH_DEAD_CLICKS, HEALTH_SLEEP_IMPR, HEALTH_MIN_ORDERS,
    BASE_COLS, METRIC_COLS, SPARKLINE_PERIODS
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

EXCLUDE_SHEETS = ['Listing', 'Portfolio ID', 'Sheet1', 'Sheet2']


# ─── File Discovery ──────────────────────────────────────────────────────────

def find_latest_updated_file(data_dir: str, pattern: str = "PPC_*_UPDATED*.xlsx") -> str | None:
    """Tìm file PPC_*_UPDATED*.xlsx mới nhất trong thư mục."""
    files = glob.glob(os.path.join(data_dir, pattern))
    if not files:
        logging.error(f"Không tìm thấy file nào khớp pattern '{pattern}' tại: {data_dir}")
        return None
    latest = max(files, key=os.path.getmtime)
    logging.info(f"File mới nhất: {os.path.basename(latest)}")
    return latest


# ─── ETL Core ────────────────────────────────────────────────────────────────

def _parse_date_blocks_from_row2(df_raw: pd.DataFrame) -> list[str]:
    """
    Trích xuất danh sách date block ("XXXX-YYYY") từ các cột merged header Row 2.
    Ví dụ cột "Ngày 3004-0405" → "3004-0405"
    """
    date_blocks = []
    seen = set()
    for col in df_raw.columns:
        col_str = str(col)
        m = re.search(r'Ng[àa]y\s*([\d\-]+)', col_str, re.IGNORECASE)
        if m:
            db = m.group(1).strip()
            if db not in seen:
                date_blocks.append(db)
                seen.add(db)
    return date_blocks


def load_ppc_updated(file_path: str) -> dict:
    """
    Đọc file PPC_*_UPDATED.xlsx và trả về dict:
    {
        'df_kw':         DataFrame tất cả keywords (long format, cột metric = kỳ gần nhất),
        'df_kw_ts':      DataFrame wide format với tất cả metric của tất cả date blocks,
        'date_blocks':   list[str] tất cả date blocks tìm thấy,
        'skus':          list[str] tất cả SKU,
    }
    Lấy file mới nhất → đọc sheet cuối / metric của date block gần nhất làm "hiện tại".
    """
    logging.info(f"Đang đọc: {os.path.basename(file_path)}")
    xl = pd.ExcelFile(file_path, engine='openpyxl')

    all_kw_rows = []
    all_kw_ts_rows = []
    all_date_blocks = []

    for sheet_name in xl.sheet_names:
        if sheet_name in EXCLUDE_SHEETS:
            continue

        # Đọc toàn bộ sheet không header để phân tích cấu trúc
        raw = pd.read_excel(file_path, sheet_name=sheet_name,
                            engine='openpyxl', header=None)
        if raw.empty or raw.shape[0] < 3:
            continue

        # Row 2 (index 1) = header chính
        headers = [str(v).strip() if pd.notna(v) else '' for v in raw.iloc[1].values]

        # Tìm date blocks từ Row 1 (index 0) — merged cells "Ngày XXXX"
        row1_vals = [str(v).strip() if pd.notna(v) else '' for v in raw.iloc[0].values]
        date_blocks_in_sheet = []
        date_block_starts = {}   # date_str → col_index (0-based)
        for col_idx, val in enumerate(row1_vals):
            m = re.search(r'Ng[àa]y\s*([\d\-]+)', val, re.IGNORECASE)
            if m:
                db = m.group(1).strip()
                if db not in date_block_starts:
                    date_blocks_in_sheet.append(db)
                    date_block_starts[db] = col_idx

        # Cập nhật toàn bộ date blocks
        for db in date_blocks_in_sheet:
            if db not in all_date_blocks:
                all_date_blocks.append(db)

        # Data rows = từ index 2 trở đi
        data_raw = raw.iloc[2:].copy()
        data_raw.columns = headers
        data_raw = data_raw.reset_index(drop=True)

        # Lấy metric cột gần nhất (date block cuối cùng trong sheet)
        latest_db = date_blocks_in_sheet[-1] if date_blocks_in_sheet else None

        for _, row in data_raw.iterrows():
            camp_name = str(row.get('Campaign Name', '')).strip()
            target    = str(row.get('Target', '')).strip()
            if not camp_name or camp_name in ('nan', '', 'Campaign Name'):
                continue

            # Base info
            base = {
                'SKU':           sheet_name,
                'Campaign Name': camp_name,
                'Loại Campaign': str(row.get('Loại Campaign', '')).strip(),
                'Target':        target,
                'Ghi chú':       str(row.get('Ghi chú', '') or row.get('Ghi chu', '')).strip(),
                'Status':        str(row.get('Status', '')).strip(),
            }

            # Match type từ Ghi chú: [[Exact]], [[Phrase]], [[Broad]]
            note = base['Ghi chú']
            mt_match = re.search(r'\[\[(Exact|Phrase|Broad|targeting)\]', note, re.IGNORECASE)
            base['Match Type'] = mt_match.group(1).capitalize() if mt_match else ''

            # Metric của date block mới nhất → dùng làm "hiện tại"
            if latest_db and latest_db in date_block_starts:
                start_col = date_block_starts[latest_db]
                metric_names = ['Impressions', 'Clicks', 'Click-through Rate', 'Spend',
                                'Sales', 'Orders', 'Units', 'Conversion Rate', 'ACOS', 'CPC', 'ROAS']
                for i, mn in enumerate(metric_names):
                    col_h = headers[start_col + i] if (start_col + i) < len(headers) else ''
                    raw_val = row.iloc[start_col + i] if (start_col + i) < len(row) else 0
                    try:
                        base[mn] = float(raw_val) if pd.notna(raw_val) and raw_val != '' else 0.0
                    except (ValueError, TypeError):
                        base[mn] = 0.0

            # Health Tag & Suggestion
            acos    = base.get('ACOS', 0) or 0
            orders  = base.get('Orders', 0) or 0
            clicks  = base.get('Clicks', 0) or 0
            impr    = base.get('Impressions', 0) or 0
            base['Health Tag']      = get_health_tag(acos, orders, clicks, impr)
            base['Suggested Action'] = get_keyword_suggestion(orders, acos, clicks)

            all_kw_rows.append(base)

            # Time Series (wide): thêm metric của TẤT CẢ date blocks
            ts_row = {k: base[k] for k in ['SKU', 'Campaign Name', 'Target', 'Match Type', 'Status']}
            for db in date_blocks_in_sheet:
                db_start = date_block_starts[db]
                for i, mn in enumerate(['Spend', 'Sales', 'Orders', 'Impressions', 'Clicks']):
                    raw_val = row.iloc[db_start + ['Impressions', 'Clicks', 'Click-through Rate',
                                                    'Spend', 'Sales', 'Orders', 'Units',
                                                    'Conversion Rate', 'ACOS', 'CPC', 'ROAS'].index(mn)
                                       + db_start] if False else 0
                    # Lấy đúng vị trí từ headers
                    metric_order = ['Impressions', 'Clicks', 'Click-through Rate', 'Spend',
                                    'Sales', 'Orders', 'Units', 'Conversion Rate', 'ACOS', 'CPC', 'ROAS']
                    if mn in metric_order:
                        offset = metric_order.index(mn)
                        col_i = db_start + offset
                        rv = row.iloc[col_i] if col_i < len(row) else 0
                        try:
                            ts_row[f'{mn}_{db}'] = float(rv) if pd.notna(rv) and rv != '' else 0.0
                        except (ValueError, TypeError):
                            ts_row[f'{mn}_{db}'] = 0.0

            all_kw_ts_rows.append(ts_row)

    df_kw    = pd.DataFrame(all_kw_rows)
    df_kw_ts = pd.DataFrame(all_kw_ts_rows)

    # Chuẩn hoá numeric
    num_cols = ['Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'Units',
                'Click-through Rate', 'Conversion Rate', 'ACOS', 'CPC', 'ROAS']
    for c in num_cols:
        if c in df_kw.columns:
            df_kw[c] = pd.to_numeric(df_kw[c], errors='coerce').fillna(0)

    skus = sorted(df_kw['SKU'].unique().tolist()) if not df_kw.empty else []

    return {
        'df_kw':      df_kw,
        'df_kw_ts':   df_kw_ts,
        'date_blocks': all_date_blocks,
        'skus':       skus,
    }


def aggregate_campaigns(df_kw: pd.DataFrame) -> pd.DataFrame:
    """
    Tổng hợp metrics từ keyword level lên campaign level bằng groupby.
    Trả về df_camp với các cột: SKU, Campaign Name, Loại Campaign, Status,
    Spend, Sales, Orders, Impressions, Clicks, ACOS, ROAS, CTR, CVR, Health Tag.
    """
    if df_kw.empty:
        return pd.DataFrame()

    grp = df_kw.groupby(['SKU', 'Campaign Name', 'Loại Campaign', 'Status'], as_index=False).agg(
        Spend=('Spend', 'sum'),
        Sales=('Sales', 'sum'),
        Orders=('Orders', 'sum'),
        Impressions=('Impressions', 'sum'),
        Clicks=('Clicks', 'sum'),
        Units=('Units', 'sum'),
    )

    # Tính các KPI phái sinh
    grp['ACOS'] = grp.apply(
        lambda r: r['Spend'] / r['Sales'] if r['Sales'] > 0 else 0, axis=1
    )
    grp['ROAS'] = grp.apply(
        lambda r: r['Sales'] / r['Spend'] if r['Spend'] > 0 else 0, axis=1
    )
    grp['CTR'] = grp.apply(
        lambda r: r['Clicks'] / r['Impressions'] if r['Impressions'] > 0 else 0, axis=1
    )
    grp['CVR'] = grp.apply(
        lambda r: r['Orders'] / r['Clicks'] if r['Clicks'] > 0 else 0, axis=1
    )
    grp['CPC'] = grp.apply(
        lambda r: r['Spend'] / r['Clicks'] if r['Clicks'] > 0 else 0, axis=1
    )

    # Health Tag
    grp['Health Tag'] = grp.apply(
        lambda r: get_health_tag(r['ACOS'], r['Orders'], r['Clicks'], r['Impressions']),
        axis=1
    )

    return grp.sort_values('Spend', ascending=False).reset_index(drop=True)


# ─── Business Logic ──────────────────────────────────────────────────────────

def get_health_tag(acos: float, orders: float, clicks: float, impressions: float) -> str:
    """Gán Health Tag cho campaign/keyword dựa trên metrics."""
    acos   = float(acos or 0)
    orders = float(orders or 0)
    clicks = float(clicks or 0)
    impr   = float(impressions or 0)

    if acos <= HEALTH_ACOS_STAR and orders >= HEALTH_MIN_ORDERS:
        return '⭐ Star'
    if acos > HEALTH_ACOS_BLEEDER and orders >= HEALTH_MIN_ORDERS:
        return '🔴 Bleeder'
    if HEALTH_ACOS_STAR < acos <= HEALTH_ACOS_BLEEDER and orders >= HEALTH_MIN_ORDERS:
        return '⚠️ Watch'
    if orders == 0 and clicks >= HEALTH_DEAD_CLICKS:
        return '💀 Dead'
    if impr < HEALTH_SLEEP_IMPR:
        return '😴 Sleep'
    return '🆕 New'


def get_keyword_suggestion(orders: float, acos: float, clicks: float) -> str:
    """Gợi ý hành động cho từ khoá."""
    orders = float(orders or 0)
    acos   = float(acos or 0)
    clicks = float(clicks or 0)

    if orders >= 2 and acos <= 0.30:
        return '📈 Tăng Bid'
    if orders == 0 and clicks >= HEALTH_DEAD_CLICKS:
        return '⏸️ Tạm dừng'
    if orders >= 1 and acos > 0.40:
        return '📉 Giảm Bid'
    return '👀 Theo dõi'


def aggregate_campaign_ts(df_kw_ts: pd.DataFrame, date_blocks: list) -> pd.DataFrame:
    """
    Tổng hợp time-series Spend và Sales theo từng date block cho mỗi campaign.
    Dùng để vẽ sparklines trong CAMPAIGNS sheet.
    Trả về DataFrame với index=Campaign Name, các cột TS_Spend_XXXX, TS_Sales_XXXX
    """
    if df_kw_ts.empty:
        return pd.DataFrame()

    result_rows = []
    for camp_name, grp in df_kw_ts.groupby('Campaign Name'):
        row = {'Campaign Name': camp_name}
        for db in date_blocks:
            spend_col = f'Spend_{db}'
            sales_col = f'Sales_{db}'
            row[f'TS_Spend_{db}'] = grp[spend_col].sum() if spend_col in grp.columns else 0
            row[f'TS_Sales_{db}'] = grp[sales_col].sum() if sales_col in grp.columns else 0
        result_rows.append(row)

    return pd.DataFrame(result_rows).set_index('Campaign Name')
