"""
fix_uploaded_campaign_names.py  —  Core Logic Layer
====================================================
Role  : Reads the Bulk JSON, detects campaigns whose names contain underscore
        separators where spaces should be (formatting bug from Amazon upload),
        and produces per-SKU rename-update JSON/xlsx files.

Input  : data/working_json/BulkSheetExport_*.json
Output : data/working_json/RENAME_UPDATE_<SKU>_*.json  (one per SKU)
         data/final_xlsx/RENAME_UPDATE_<SKU>_*.xlsx    (via excel_json_handler)

Zero Excel I/O — pure dict operations.
"""

import os
import sys
import json
import glob
import re
import logging

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, 'io_handlers'))
from excel_json_handler import json_to_xlsx, AMAZON_BULK_COLUMNS

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

JSON_DIR  = os.path.join(BASE_DIR, 'data', 'working_json')
FINAL_DIR = os.path.join(BASE_DIR, 'data', 'final_xlsx')
os.makedirs(FINAL_DIR, exist_ok=True)


# ─── Name parser (unchanged business logic) ──────────────────────────────────

def parse_name(name: str):
    """
    Detect and fix campaign/ad group names that use underscores instead of
    spaces in the keyword portion.
    Returns (fixed_name, sku) or (original_name, 'UNKNOWN') if no fix needed.
    """
    if not isinstance(name, str):
        return name, 'UNKNOWN'

    match = re.search(r'_(KT|PT|AU)_', name)
    if not match:
        return name, 'UNKNOWN'

    type_code = match.group(1)
    idx       = match.start()
    sku       = name[:idx]
    rest      = name[match.end():]
    parts     = rest.split('_')

    if len(parts) < 2:
        return name, sku

    match_type_part = parts[0]
    placement_parts = []
    keyword_parts   = []

    for i in range(len(parts) - 1, 0, -1):
        if re.match(r'^\d+[TRP]*$', parts[i]):
            placement_parts.insert(0, parts[i])
        else:
            keyword_parts = parts[1:i + 1]
            break

    if not keyword_parts:
        return name, sku

    fixed_keyword = ' '.join(keyword_parts)
    placement_str = '_'.join(placement_parts)

    if placement_str:
        return f"{sku}_{type_code}_{match_type_part}_{fixed_keyword}_{placement_str}", sku
    else:
        return f"{sku}_{type_code}_{match_type_part}_{fixed_keyword}", sku


def _safe_id(val: str) -> str:
    val = str(val).strip().rstrip('.0')
    return '' if val.lower() in ('nan', 'none', '') else val


# ─── Core processing ─────────────────────────────────────────────────────────

def detect_rename_candidates(records: list) -> dict:
    """
    Scan all Campaign and Ad Group rows for naming issues.
    Returns sku_buckets: {sku: [rename_row_dict, ...]}
    """
    sku_buckets = {}

    for row in records:
        entity = str(row.get('Entity', '')).strip().lower()
        state  = str(row.get('State', 'Enabled')).strip().title() or 'Enabled'

        if entity == 'campaign':
            old_name = str(row.get('Campaign Name', '')).strip()
            new_name, sku = parse_name(old_name)
            if old_name and old_name != new_name:
                rename_row = {
                    'Product':    'Sponsored Products',
                    'Entity':     'Campaign',
                    'Operation':  'Update',
                    'Campaign ID': _safe_id(str(row.get('Campaign ID', ''))),
                    'Campaign Name': new_name,
                    'State': state
                }
                sku_buckets.setdefault(sku, []).append(rename_row)

        elif entity == 'ad group':
            old_name = str(row.get('Ad Group Name', '')).strip()
            new_name, sku = parse_name(old_name)
            if old_name and old_name != new_name:
                rename_row = {
                    'Product':     'Sponsored Products',
                    'Entity':      'Ad Group',
                    'Operation':   'Update',
                    'Campaign ID':  _safe_id(str(row.get('Campaign ID', ''))),
                    'Ad Group ID':  _safe_id(str(row.get('Ad Group ID', ''))),
                    'Ad Group Name': new_name,
                    'State': state
                }
                sku_buckets.setdefault(sku, []).append(rename_row)

    return sku_buckets


# ─── Entry point ─────────────────────────────────────────────────────────────

def main():
    print("--- FIX UPLOADED CAMPAIGN NAMES (JSON Mode) ---")

    bulk_jsons = sorted(glob.glob(os.path.join(JSON_DIR, 'BulkSheetExport_*.json')))
    if not bulk_jsons:
        logging.error("No BulkSheetExport_*.json found. Run rebuild_internal_file.py first.")
        return

    for bj in bulk_jsons:
        basename = os.path.basename(bj).replace('.json', '')
        logging.info(f"Processing: {os.path.basename(bj)}")

        with open(bj, 'r', encoding='utf-8') as f:
            records = json.load(f)

        sku_buckets = detect_rename_candidates(records)

        if not sku_buckets:
            logging.info("  → No campaigns with naming issues found.")
            continue

        total = sum(len(v) for v in sku_buckets.values())
        logging.info(f"  → {total} rename actions across {len(sku_buckets)} SKUs.")

        for sku, update_rows in sku_buckets.items():
            safe_sku = re.sub(r'[^\w\-]', '_', sku)
            out_name  = f"RENAME_UPDATE_{safe_sku}_{basename}"

            # Write JSON
            out_json = os.path.join(JSON_DIR, out_name + '.json')
            with open(out_json, 'w', encoding='utf-8') as f:
                json.dump(update_rows, f, ensure_ascii=False, indent=2)

            # Write xlsx (Text-formatted for Amazon upload)
            out_xlsx = os.path.join(FINAL_DIR, out_name + '.xlsx')
            ok = json_to_xlsx(out_json, out_xlsx)

            status = '[OK]' if ok else '[FAIL]'
            print(f"  {status} SKU '{sku}': {len(update_rows)} operations → {out_name}.xlsx")

    logging.info("FIX NAMES COMPLETE.")


if __name__ == "__main__":
    main()
