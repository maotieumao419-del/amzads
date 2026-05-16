"""
rebuild_internal_file.py  —  I/O Boundary Layer
================================================
Role  : Clean & normalise raw xlsx files from data/raw_xlsx/.
Output: 1) Formatted xlsx reference copies (optional, for human inspection)
        2) Canonical JSON files in data/working_json/ for all downstream scripts.

Execution order (mandatory):
  python io_handlers/rebuild_internal_file.py
  → then run any core_logic/*.py script
"""

import os
import sys
import glob
import logging
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Allow import of sibling io_handlers module
BASE_DIR   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, 'io_handlers'))
from excel_json_handler import xlsx_to_json, internal_xlsx_to_json

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

RAW_DIR  = os.path.join(BASE_DIR, 'data', 'raw_xlsx')
JSON_DIR = os.path.join(BASE_DIR, 'data', 'working_json')
os.makedirs(JSON_DIR, exist_ok=True)

_THIN = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'),  bottom=Side(style='thin')
)


# ─── Helpers ────────────────────────────────────────────────────────────────

def _clean_workbook_from_df(df: pd.DataFrame, sheet_name: str,
                            is_bulk: bool = False) -> openpyxl.Workbook:
    """
    Build a clean openpyxl Workbook from a DataFrame.
    Applies minimal but correct formatting (bold headers, thin borders).
    Returns the Workbook object (NOT saved — caller decides where to save).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    df.fillna('', inplace=True)
    # Write header
    for c_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.font   = Font(bold=True)
        cell.border = _THIN
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Write data
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx + 2, column=c_idx, value=val if val != '' else None)
            cell.border = _THIN
            if not is_bulk:
                cell.alignment = Alignment(vertical='center', wrap_text=True)

    # Column widths
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        sample = [c.value for c in col_cells[:50] if c.value]
        max_len = max((len(str(v)) for v in sample), default=8)
        ws.column_dimensions[col_letter].width = min(max_len + 2, 30 if is_bulk else 50)

    ws.row_dimensions[1].height = 18
    return wb


# ─── Core rebuild functions ──────────────────────────────────────────────────

def rebuild_bulk(src_path: str) -> bool:
    """
    Process an Amazon Bulk Master xlsx:
      1. Read with pandas (preserves dtype).
      2. Convert directly to working_json/BulkSheetExport_*.json
    No xlsx copy needed — JSON is the canonical representation.
    """
    filename = os.path.basename(src_path)
    json_out = os.path.join(JSON_DIR, filename.replace('.xlsx', '.json'))

    logging.info(f"[rebuild_bulk] {filename}")
    ok = xlsx_to_json(src_path, json_out, sheet_name=None)
    if ok:
        logging.info(f"  → JSON: {os.path.basename(json_out)}")
    return ok


def rebuild_internal(src_path: str) -> bool:
    """
    Process a PPC Internal xlsx:
      1. Load with openpyxl (to honour the multi-sheet, merged-cell structure).
      2. Convert to working_json/PPC_*.json  using internal_xlsx_to_json().
    """
    filename = os.path.basename(src_path)
    json_out = os.path.join(JSON_DIR, filename.replace('.xlsx', '.json'))

    logging.info(f"[rebuild_internal] {filename}")
    ok = internal_xlsx_to_json(src_path, json_out)
    if ok:
        logging.info(f"  → JSON: {os.path.basename(json_out)}")
    return ok


# ─── Entry point ────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 60)
    print("  REBUILD — Convert raw_xlsx → working_json")
    print("=" * 60)

    # ── 1. Amazon Bulk files ─────────────────────────────────────────
    bulk_files = sorted(
        [f for f in glob.glob(os.path.join(RAW_DIR, 'BulkSheetExport_*.xlsx'))
         if not os.path.basename(f).startswith('~$')],
        key=os.path.getmtime, reverse=True
    )
    if not bulk_files:
        logging.warning(f"No BulkSheetExport_*.xlsx found in {RAW_DIR}")
    else:
        for src in bulk_files:
            ok = rebuild_bulk(src)
            status = "[OK]" if ok else "[FAIL]"
            print(f"  {status} Bulk  : {os.path.basename(src)}")

    # ── 2. PPC Internal files ────────────────────────────────────────
    ppc_files = sorted(
        [f for f in glob.glob(os.path.join(RAW_DIR, 'PPC_*.xlsx'))
         if '_UPDATED' not in f and not os.path.basename(f).startswith('~$')],
        key=os.path.getmtime, reverse=True
    )
    if not ppc_files:
        logging.warning(f"No PPC_*.xlsx found in {RAW_DIR}")
    else:
        for src in ppc_files:
            ok = rebuild_internal(src)
            status = "[OK]" if ok else "[FAIL]"
            print(f"  {status} Internal: {os.path.basename(src)}")

    print(f"\n  JSON files ready in: data/working_json/")
    print("  → Next: run core_logic scripts")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
