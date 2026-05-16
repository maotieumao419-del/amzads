"""
excel_json_handler.py  —  I/O Boundary Layer
=============================================
Central converter between .xlsx files and .json intermediates.

Functions
---------
  xlsx_to_json()           : Amazon Bulk xlsx  → flat JSON (List[Dict])
  json_to_xlsx()           : flat JSON         → Amazon-formatted xlsx
  internal_xlsx_to_json()  : PPC Internal xlsx → structured JSON (Dict)
  json_to_internal_xlsx()  : structured JSON   → PPC Internal xlsx
"""

import os
import json
import logging
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ─── Constants ──────────────────────────────────────────────────────────────

AMAZON_BULK_COLUMNS = [
    'Product', 'Entity', 'Operation', 'Campaign ID', 'Ad Group ID', 'Portfolio ID',
    'Ad ID', 'Keyword ID', 'Product Targeting ID', 'Campaign Name', 'Ad Group Name',
    'Start Date', 'End Date', 'Targeting Type', 'State', 'Daily Budget', 'SKU',
    'Ad Group Default Bid', 'Bid', 'Keyword Text', 'Native Language Keyword',
    'Native Language Locale', 'Match Type', 'Bidding Strategy', 'Placement',
    'Percentage', 'Product Targeting Expression', 'Audience ID',
    'Shopper Cohort Percentage', 'Shopper Cohort Type'
]

METRIC_COLUMNS = [
    'Impressions', 'Clicks', 'Click-through Rate', 'Spend', 'Sales',
    'Orders', 'Units', 'Conversion Rate', 'ACOS', 'CPC', 'ROAS'
]

BASE_INTERNAL_HEADERS = ['STT', 'Campaign Name', 'Loại Campaign', 'Target', 'Ghi chú', 'Status']

# Columns that must be cast to clean strings (no scientific notation, no trailing .0)
_ID_COLUMNS = {
    'Campaign ID', 'Ad Group ID', 'Portfolio ID', 'Ad ID',
    'Keyword ID', 'Product Targeting ID', 'Audience ID'
}
# Columns that must be cast to float (0.0 if blank/invalid)
_FLOAT_COLUMNS = {
    'Bid', 'Daily Budget', 'Ad Group Default Bid', 'Percentage',
    'Shopper Cohort Percentage', 'Impressions', 'Clicks', 'Spend',
    'Sales', 'Orders', 'Units', 'Click-through Rate',
    'Conversion Rate', 'ACOS', 'CPC', 'ROAS'
}

_THIN = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'),  bottom=Side(style='thin')
)


# ─── Dtype enforcement ───────────────────────────────────────────────────────

def enforce_dtypes(record: dict) -> dict:
    """
    Enforce strict data types on a single row dict:
      - ID columns  : clean integer-string (strip trailing .0 and spaces)
      - Float cols  : Python float, 0.0 if blank / unparseable
      - Everything else : stripped string, None if truly empty
    """
    out = {}
    for k, v in record.items():
        # ── ID columns ────────────────────────────────────────────────
        if k in _ID_COLUMNS:
            s = str(v).strip() if v is not None else ''
            if s in ('', 'nan', 'None', 'NaN'):
                out[k] = ''
            else:
                try:
                    out[k] = str(int(float(s)))   # '123456789.0' → '123456789'
                except (ValueError, TypeError):
                    out[k] = s

        # ── Float / numeric columns ───────────────────────────────────
        elif k in _FLOAT_COLUMNS:
            if v is None or str(v).strip() in ('', 'nan', 'None', 'NaN'):
                out[k] = 0.0
            else:
                try:
                    out[k] = float(v)
                except (ValueError, TypeError):
                    out[k] = 0.0

        # ── All other columns (text) ──────────────────────────────────
        else:
            s = str(v).strip() if v is not None else ''
            out[k] = None if s in ('nan', 'None', 'NaN') else (s or None)

    return out



# ═══════════════════════════════════════════════════════════════════════════
# 1.  Amazon Bulk  xlsx → flat JSON  (strict dtypes)
# ═══════════════════════════════════════════════════════════════════════════

def xlsx_to_json(xlsx_path: str, json_path: str, sheet_name: str = None) -> bool:
    """
    Read one sheet from an Amazon Bulk xlsx and write a strictly-typed flat JSON.

    Data-type guarantees
    --------------------
    - ID columns (Campaign ID, Keyword ID …) → clean integer-strings, no scientific notation.
    - Numeric columns (Bid, Budget, Spend …)  → Python float, 0.0 if blank.
    - Text columns                             → stripped string, None if empty.
    - NaN / blank cells                        → None (JSON null) for text, 0.0 for numbers.

    Output: List[Dict] serialised to json_path.
    """
    try:
        xls = pd.ExcelFile(xlsx_path, engine='openpyxl')
        if sheet_name is None:
            sheet_name = next(
                (s for s in xls.sheet_names if 'Sponsored Product' in s),
                xls.sheet_names[0]
            )

        # Read everything as object first so we control casting
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name,
                           dtype=object, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        # Apply strict dtype enforcement row-by-row
        raw_records = df.where(pd.notna(df), None).to_dict(orient='records')
        records = [enforce_dtypes(r) for r in raw_records]

        os.makedirs(os.path.dirname(json_path), exist_ok=True)
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=2)

        logging.info(f"[xlsx_to_json] {len(records)} rows → {os.path.basename(json_path)}")
        return True

    except Exception as e:
        logging.error(f"[xlsx_to_json] Failed for {xlsx_path}: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════════════
# 2.  Flat JSON → Amazon Bulk xlsx
# ═══════════════════════════════════════════════════════════════════════════

def json_to_xlsx(json_path: str, xlsx_path: str,
                 sheet_name: str = 'Sponsored Products Campaigns',
                 column_template: list = None) -> bool:
    """
    Write a flat JSON (List[Dict]) to an Amazon-formatted xlsx.

    Guarantees:
    - All columns forced to Text format (prevents ID scientific-notation corruption).
    - Columns ordered by column_template (defaults to AMAZON_BULK_COLUMNS).
    - Missing template columns are added as empty strings.
    """
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            records = json.load(f)

        columns = column_template or AMAZON_BULK_COLUMNS
        df = pd.DataFrame(records)
        for col in columns:
            if col not in df.columns:
                df[col] = ''
        df = df.reindex(columns=columns, fill_value='')

        os.makedirs(os.path.dirname(xlsx_path), exist_ok=True)
        with pd.ExcelWriter(xlsx_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            workbook  = writer.book
            worksheet = writer.sheets[sheet_name]
            text_fmt  = workbook.add_format({'num_format': '@'})
            hdr_fmt   = workbook.add_format({'num_format': '@', 'bold': True})
            for col_num in range(len(columns)):
                worksheet.set_column(col_num, col_num, 22, text_fmt)
            for col_num, _ in enumerate(columns):
                worksheet.write(0, col_num, columns[col_num], hdr_fmt)

        logging.info(f"[json_to_xlsx] {len(records)} rows → {os.path.basename(xlsx_path)}")
        return True

    except Exception as e:
        logging.error(f"[json_to_xlsx] Failed for {json_path}: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════════════
# 3.  PPC Internal xlsx → structured JSON
# ═══════════════════════════════════════════════════════════════════════════

def internal_xlsx_to_json(xlsx_path: str, json_path: str) -> bool:
    """
    Read a multi-sheet PPC Internal xlsx and produce a structured JSON.

    Output JSON schema
    ------------------
    {
      "Listing": [ {col: val, ...}, ... ],

      "<SKU_name>": {
        "sku_title":    "SKU: <name>",
        "base_headers": ["STT", "Campaign Name", "Loại Campaign", "Target", "Ghi chú"],
        "date_blocks":  ["3004-0405", ...],
        "rows": [
          {
            "STT": 1,
            "Campaign Name": "...",
            "Loại Campaign": "Keyword (Từ khóa)",
            "Target": "nurse gifts",
            "Ghi chú": "[[Exact]-[0.56]-[10T-0R-0P]]",
            "metrics_by_date": {
              "3004-0405": {
                "Impressions": 109, "Clicks": 0, ..., "Placement Note": ""
              }
            }
          },
          ...
        ]
      }
    }
    """
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        output = {}

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # ── Detect Sheet Type ────────────────────────────────────────
            val_a1 = str(ws.cell(row=1, column=1).value or '').strip()
            val_b2 = str(ws.cell(row=2, column=2).value or '').strip()
            val_d2 = str(ws.cell(row=2, column=4).value or '').strip()
            
            # It's an SKU sheet if A1 starts with SKU: OR Row 2 has standard headers
            is_sku_sheet = val_a1.startswith('SKU:') or (val_b2 == 'Campaign Name' and val_d2 == 'Target')

            if sheet_name == 'Listing':
                headers = [
                    str(ws.cell(row=1, column=c).value or '').strip()
                    for c in range(1, ws.max_column + 1)
                ]
                rows = []
                for r in range(2, ws.max_row + 1):
                    row_dict = {}
                    for c, h in enumerate(headers, 1):
                        val = ws.cell(row=r, column=c).value
                        row_dict[h] = str(val).strip() if val is not None else ''
                    if any(v for v in row_dict.values()):
                        rows.append(row_dict)
                output['Listing'] = rows
                continue
            
            if sheet_name == 'Portfolio ID' or not is_sku_sheet:
                # Simple list of lists for non-SKU sheets
                table = []
                for r in range(1, ws.max_row + 1):
                    row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
                    if any(v is not None and str(v).strip() for v in row_vals):
                        table.append(row_vals)
                output[sheet_name] = table
                continue

            # ── SKU sheet ────────────────────────────────────────────────
            sku_title = str(ws.cell(row=1, column=1).value or '').strip()

            # Row 2: all column headers (base + date-block sub-headers)
            row2 = [
                str(ws.cell(row=2, column=c).value or '').strip()
                for c in range(1, ws.max_column + 1)
            ]

            # Detect where base columns end (first metric column marks boundary)
            base_col_count = 6  # default
            for i, h in enumerate(row2):
                if h in METRIC_COLUMNS:
                    base_col_count = i
                    break
            base_headers = [h for h in row2[:base_col_count] if h]
            # Standardize: Ensure at least 6 base columns, 6th one is 'Status'
            if len(base_headers) < 6:
                base_headers.extend(['Status'] * (6 - len(base_headers)))
            else:
                base_headers[5] = 'Status'

            # Detect date blocks from Row-1 merged cells
            date_blocks = []
            date_block_cols = {}   # date_str → start_col (1-indexed)
            for merged in ws.merged_cells.ranges:
                if merged.min_row == 1 and merged.min_col > base_col_count:
                    raw = str(ws.cell(row=1, column=merged.min_col).value or '').strip()
                    date_str = raw.replace('Ngày ', '').replace('Ngay ', '').strip()
                    if date_str:
                        date_blocks.append(date_str)
                        date_block_cols[date_str] = merged.min_col

            # Parse data rows
            rows = []
            for r in range(3, ws.max_row + 1):
                row_dict = {}
                has_data = False
                for c_idx, h in enumerate(base_headers, 1):
                    val = ws.cell(row=r, column=c_idx).value
                    row_dict[h] = val if val is not None else ''
                    if val and str(val).strip():
                        has_data = True
                if not has_data:
                    continue

                # Metrics per date block
                row_dict['metrics_by_date'] = {}
                for date_str, start_col in date_block_cols.items():
                    metrics = {}
                    for i, m in enumerate(METRIC_COLUMNS):
                        val = ws.cell(row=r, column=start_col + i).value
                        metrics[m] = val if val is not None else 0
                    note_val = ws.cell(row=r, column=start_col + 11).value
                    metrics['Placement Note'] = str(note_val or '').strip()
                    row_dict['metrics_by_date'][date_str] = metrics

                rows.append(row_dict)

            output[sheet_name] = {
                'sku_title':    sku_title,
                'base_headers': base_headers,
                'date_blocks':  date_blocks,
                'rows':         rows
            }

        os.makedirs(os.path.dirname(json_path), exist_ok=True)
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False, indent=2, default=str)

        logging.info(
            f"[internal_xlsx_to_json] {len(wb.sheetnames)} sheets → {os.path.basename(json_path)}"
        )
        return True

    except Exception as e:
        logging.error(f"[internal_xlsx_to_json] Failed for {xlsx_path}: {e}")
        import traceback; traceback.print_exc()
        return False


# ═══════════════════════════════════════════════════════════════════════════
# 4.  Structured JSON → PPC Internal xlsx
# ═══════════════════════════════════════════════════════════════════════════

def json_to_internal_xlsx(json_path: str, xlsx_path: str) -> bool:
    """
    Convert a structured PPC Internal JSON back to a multi-sheet xlsx.

    Formatting applied
    ------------------
    - SKU sheets  Row 1 : Merged title cell, bold, centred.
    - SKU sheets  Row 2 : Base headers + date-block sub-headers, bold, bordered.
    - Date blocks Row 1 : Merged "Ngày XXXX" cell above its 12 sub-columns.
    - Data rows          : Thin border on all occupied cells.
    - Listing sheet      : Simple tabular format.
    """
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        for sheet_name, sheet_data in data.items():
            ws = wb.create_sheet(title=sheet_name[:31])

            # ── Tabular sheets ────────────────────────────────────────────
            if isinstance(sheet_data, list) and sheet_data:
                # Case A: List of lists (e.g. Portfolio ID)
                if isinstance(sheet_data[0], list):
                    for r_idx, row_vals in enumerate(sheet_data, 1):
                        for c_idx, val in enumerate(row_vals, 1):
                            cell = ws.cell(row=r_idx, column=c_idx, value=val)
                            if r_idx == 1:
                                cell.font = Font(bold=True)
                                cell.border = _THIN
                    # Auto-fit columns for tabular sheets (Portfolio ID)
                    for col_cells in ws.columns:
                        col_letter = get_column_letter(col_cells[0].column)
                        max_len = max((len(str(c.value or '')) for c in col_cells), default=8)
                        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)
                    for row_cells in ws.iter_rows():
                        ws.row_dimensions[row_cells[0].row].height = 20
                    continue

                # Case B: List of dicts (e.g. Listing)
                raw_headers = list(sheet_data[0].keys())
                # Reference has '  STT' with double space in Listing
                headers = [h if h != 'STT' else '  STT' for h in raw_headers]

                for c_idx, h in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=c_idx, value=h)
                    cell.font  = Font(bold=True)
                    cell.border = _THIN

                for r_idx, row_dict in enumerate(sheet_data, 2):
                    for c_idx, h in enumerate(raw_headers, 1):
                        val = row_dict.get(h, '')
                        # Reference has '\n' prefix in SKU column of Listing
                        if sheet_name == 'Listing' and h == 'SKU' and val and not str(val).startswith('\n'):
                            val = f"\n{val}"
                        cell = ws.cell(row=r_idx, column=c_idx, value=val)
                        cell.alignment = Alignment(wrap_text=True, vertical='top')

                # Auto-fit columns for Listing sheet
                for col_cells in ws.columns:
                    col_letter = get_column_letter(col_cells[0].column)
                    max_len = max((len(str(c.value or '').replace('\n', '')) for c in col_cells), default=8)
                    ws.column_dimensions[col_letter].width = min(max_len + 4, 60)

                # Auto-fit row height for Listing (based on newline count)
                for row_cells in ws.iter_rows():
                    r = row_cells[0].row
                    max_lines = max(
                        (str(c.value or '').count('\n') + 1 for c in row_cells),
                        default=1
                    )
                    ws.row_dimensions[r].height = max(20, max_lines * 15)
                continue

            # ── SKU sheet ────────────────────────────────────────────────
            sku_title    = sheet_data.get('sku_title', f'SKU: {sheet_name}')
            base_headers = sheet_data.get('base_headers', BASE_INTERNAL_HEADERS)
            date_blocks  = sheet_data.get('date_blocks', [])
            rows         = sheet_data.get('rows', [])

            base_col_count = len(base_headers)
            sub_headers = METRIC_COLUMNS + ['Placement Note']

            # ── Row 1: SKU title (merges only over the base columns) ──────
            title_end = max(base_col_count, 6)
            title_cell = ws.cell(row=1, column=1, value=sku_title)
            title_cell.font      = Font(bold=True, size=13)
            title_cell.alignment = Alignment(horizontal='center', vertical='center')
            if title_end > 1:
                ws.merge_cells(start_row=1, start_column=1,
                               end_row=1, end_column=title_end)
            ws.row_dimensions[1].height = 22

            # ── Row 2: base column headers ────────────────────────────────
            for c_idx, h in enumerate(base_headers, 1):
                cell = ws.cell(row=2, column=c_idx, value=h)
                cell.font      = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border    = _THIN

            # ── Date blocks: Row 1 merged header + Row 2 sub-headers ─────
            current_col = title_end + 1
            date_block_start_cols = {}

            for date_str in date_blocks:
                date_block_start_cols[date_str] = current_col

                # Write value into top-left cell of this block FIRST
                dh = ws.cell(row=1, column=current_col, value=f"Ngày {date_str}")
                dh.font      = Font(bold=True)
                dh.alignment = Alignment(horizontal='center')
                # Then merge it rightward over 11 more columns (12 total)
                ws.merge_cells(start_row=1, start_column=current_col,
                               end_row=1, end_column=current_col + 11)

                # Row 2 sub-headers for this block
                for i, h in enumerate(sub_headers):
                    cell = ws.cell(row=2, column=current_col + i, value=h)
                    cell.font      = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')
                    cell.border    = _THIN

                current_col += 12

            # ── Row 3+: data ──────────────────────────────────────────────
            for r_idx, row_dict in enumerate(rows, 3):
                # Base columns
                for c_idx, h in enumerate(base_headers, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=row_dict.get(h, ''))
                    cell.border = _THIN

                # Metric columns per date block
                metrics_by_date = row_dict.get('metrics_by_date', {})
                for date_str, start_col in date_block_start_cols.items():
                    m_data = metrics_by_date.get(date_str, {})
                    for i, m in enumerate(sub_headers):
                        val = m_data.get(m, 0 if m != 'Placement Note' else '')
                        cell = ws.cell(row=r_idx, column=start_col + i, value=val)
                        cell.border = _THIN

            # Column widths
            for col_cells in ws.columns:
                col_letter = get_column_letter(col_cells[0].column)
                max_len = max(
                    (len(str(c.value or '')) for c in col_cells[:50]),
                    default=8
                )
                ws.column_dimensions[col_letter].width = min(max_len + 2, 40)

        os.makedirs(os.path.dirname(xlsx_path), exist_ok=True)
        wb.save(xlsx_path)
        logging.info(f"[json_to_internal_xlsx] → {os.path.basename(xlsx_path)}")
        return True

    except Exception as e:
        logging.error(f"[json_to_internal_xlsx] Failed: {e}")
        import traceback; traceback.print_exc()
        return False
