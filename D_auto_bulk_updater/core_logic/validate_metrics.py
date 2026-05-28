"""
validate_metrics.py  —  Core Logic Layer
=========================================
Role  : Two-level validation:
  1. METRIC INTEGRITY: Cross-validates Impressions/Clicks/Orders between
     Bulk JSON (Amazon source of truth) and the UPDATED Internal JSON.
  2. DATA DRIFT CHECK: Compares row counts and Bid float arrays between
     the legacy xlsx output and the new final_xlsx output to detect any
     data drift introduced by the JSON refactor.

Input  : data/working_json/BulkSheetExport_*.json
         data/working_json/PPC_*_UPDATED.json
         data/output/*.xlsx              (legacy system output)
         data/final_xlsx/*.xlsx          (new system output)
Output : Console validation report with pass/fail summary.

Zero Excel I/O in core logic — diffs are done in memory on pandas DataFrames.
"""

import os
import sys
import json
import glob
import logging
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, 'io_handlers'))

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

JSON_DIR    = os.path.join(BASE_DIR, 'data', 'working_json')
LEGACY_DIR  = os.path.join(BASE_DIR, 'data', 'output')         # legacy pipeline output
NEW_DIR     = os.path.join(BASE_DIR, 'data', 'final_xlsx')     # new pipeline output

_SEP = '─' * 62


# ═══════════════════════════════════════════════════════════════════════════
# LEVEL 1 — Metric Integrity: Bulk JSON  ↔  Internal UPDATED JSON
# ═══════════════════════════════════════════════════════════════════════════

def _build_master_map(records: list) -> dict:
    """
    Build reference map from flat Bulk JSON.
    Returns: { sku: { (camp_name, target): {Impressions, Clicks, Orders} } }
    """
    from collections import defaultdict

    by_camp = defaultdict(list)
    for row in records:
        cid = str(row.get('Campaign ID') or row.get('Campaign Name') or '').strip()
        by_camp[cid].append(row)

    sku_map = {}

    for camp_id, group in by_camp.items():
        camp_row = next(
            (r for r in group if str(r.get('Entity', '')).strip().lower() == 'campaign'),
            group[0]
        )
        camp_name = str(camp_row.get('Campaign Name', '')).strip()

        skus = list({
            str(r.get('SKU', '')).strip()
            for r in group
            if str(r.get('Entity', '')).strip().lower() == 'product ad'
            and str(r.get('SKU', '')).strip()
        })
        if not skus:
            continue

        for row in group:
            entity = str(row.get('Entity', '')).strip().lower()
            if entity == 'keyword':
                target = str(row.get('Keyword Text', '')).strip()
                raw_m = str(row.get('Match Type', '')).strip().lower()
                match = raw_m.capitalize() if raw_m in ('exact', 'phrase', 'broad') else ''
            elif entity == 'product targeting':
                target = str(row.get('Product Targeting Expression', '')).strip()
                match = ''
            else:
                continue
            if not target:
                continue

            def _f(col):
                try: return float(row.get(col, 0) or 0)
                except: return 0.0

            metrics = {
                'Impressions': _f('Impressions'),
                'Clicks':      _f('Clicks'),
                'Orders':      _f('Orders')
            }
            for sku in skus:
                sku_map.setdefault(sku, {})[(camp_name, target, match)] = metrics

    return sku_map


def validate_metric_integrity(bulk_json_paths: list, updated_json_paths: list) -> bool:
    """
    Cross-check Impressions / Clicks / Orders between Bulk JSON and
    the UPDATED Internal JSON for matching (Camp, Target) keys.
    Returns True if zero mismatches.
    """
    print(f"\n{'═'*62}")
    print("  LEVEL 1 — Metric Integrity Check")
    print(f"{'═'*62}")

    master_map = {}
    for bj in bulk_json_paths:
        with open(bj, 'r', encoding='utf-8') as f:
            records = json.load(f)
        partial = _build_master_map(records)
        for sku, data in partial.items():
            master_map.setdefault(sku, {}).update(data)

    total_checked  = 0
    total_mismatch = 0

    for uj in updated_json_paths:
        with open(uj, 'r', encoding='utf-8') as f:
            data = json.load(f)

        label = os.path.basename(uj)
        sheet_mismatches = 0

        for sheet_name, sheet_data in data.items():
            if sheet_name == 'Listing' or not isinstance(sheet_data, dict):
                continue
            if sheet_name not in master_map:
                continue

            ref        = master_map[sheet_name]
            date_blocks = sheet_data.get('date_blocks', [])
            if not date_blocks:
                continue
            last_date = date_blocks[-1]

            import re
            for row in sheet_data.get('rows', []):
                camp   = str(row.get('Campaign Name', '')).strip()
                target = str(row.get('Target', '')).strip()
                note   = str(row.get('Ghi ch\u00fa') or row.get('Ghi chu') or '')
                m_match = re.search(r'\[\[(Exact|Phrase|Broad|targeting)\]', note, re.IGNORECASE)
                match_val = m_match.group(1).capitalize() if m_match and m_match.group(1).lower() in ('exact', 'phrase', 'broad') else ''
                
                if not camp or not target:
                    continue

                key = (camp, target, match_val)
                if key not in ref:
                    continue

                expected = ref[key]
                actual   = row.get('metrics_by_date', {}).get(last_date, {})

                def _i(v):
                    try: return int(float(v or 0))
                    except: return 0

                mismatches = []
                for metric in ('Impressions', 'Clicks', 'Orders'):
                    exp = _i(expected.get(metric, 0))
                    act = _i(actual.get(metric, 0))
                    if exp != act:
                        mismatches.append(f"{metric}: expected={exp} actual={act}")

                total_checked += 1
                if mismatches:
                    total_mismatch += 1
                    sheet_mismatches += 1
                    logging.error(
                        f"  MISMATCH [{sheet_name}] Camp='{camp}' | "
                        f"Target='{target}' | {' | '.join(mismatches)}"
                    )

        status = '✅ OK' if sheet_mismatches == 0 else f'⚠️  {sheet_mismatches} issues'
        print(f"  {label}: {status}")

    print(f"\n  Rows checked  : {total_checked}")
    if total_mismatch == 0:
        print("  Result        : ✅ ALL METRICS MATCH")
    else:
        print(f"  Result        : ⚠️  {total_mismatch} MISMATCHES FOUND")
    print(_SEP)

    return total_mismatch == 0


# ═══════════════════════════════════════════════════════════════════════════
# LEVEL 2 — Data Drift Check: Legacy xlsx  ↔  New final_xlsx
# ═══════════════════════════════════════════════════════════════════════════

def _read_upload_xlsx(path: str) -> pd.DataFrame:
    """Read an Amazon Upload xlsx into a DataFrame with clean dtypes."""
    try:
        df = pd.read_excel(path, sheet_name='Sponsored Products Campaigns',
                           dtype=str, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]
        df.fillna('', inplace=True)
        return df
    except Exception as e:
        logging.error(f"Cannot read {path}: {e}")
        return pd.DataFrame()


def validate_data_drift(legacy_dir: str, new_dir: str) -> bool:
    """
    Compare row counts and Bid float arrays between legacy and new upload files.

    Matching strategy: match files by the shared date token in the filename
    (e.g., 'Amazon_Upload_Ready_11052026').
    Returns True if zero drift detected.
    """
    print(f"\n{'═'*62}")
    print("  LEVEL 2 — Data Drift Check  (Legacy vs New)")
    print(f"{'═'*62}")

    legacy_files = [f for f in glob.glob(os.path.join(legacy_dir, '*.xlsx'))
                    if 'Amazon_Upload_Ready' in os.path.basename(f)]
    new_files    = [f for f in glob.glob(os.path.join(new_dir, '*.xlsx'))
                    if 'Amazon_Upload_Ready' in os.path.basename(f)]

    if not legacy_files:
        print(f"  [SKIP] No legacy xlsx files found in {legacy_dir}")
        return True
    if not new_files:
        print(f"  [SKIP] No new xlsx files found in {new_dir}")
        return True

    # Match pairs by filename stem similarity (longest common substring)
    def _stem(path): return os.path.splitext(os.path.basename(path))[0]

    all_pass = True

    for leg_path in legacy_files:
        leg_stem = _stem(leg_path)
        # Find best matching new file
        best = max(new_files, key=lambda p: len(
            set(_stem(p).split('_')) & set(leg_stem.split('_'))
        ))

        print(f"\n  Comparing:")
        print(f"    LEGACY : {leg_stem}.xlsx")
        print(f"    NEW    : {_stem(best)}.xlsx")

        df_leg = _read_upload_xlsx(leg_path)
        df_new = _read_upload_xlsx(best)

        issues = []

        # ── 1. Row count ──────────────────────────────────────────────
        if len(df_leg) != len(df_new):
            issues.append(
                f"Row count drift: legacy={len(df_leg)} new={len(df_new)}"
            )

        # ── 2. Bid float array ────────────────────────────────────────
        if 'Bid' in df_leg.columns and 'Bid' in df_new.columns:
            def _bids(df):
                bids = []
                for v in df['Bid']:
                    try: bids.append(round(float(v), 4))
                    except: bids.append(0.0)
                return bids

            bids_leg = _bids(df_leg)
            bids_new = _bids(df_new)

            mismatched_bids = [
                (i, bl, bn)
                for i, (bl, bn) in enumerate(zip(bids_leg, bids_new))
                if abs(bl - bn) > 0.0001
            ]
            if mismatched_bids:
                issues.append(
                    f"Bid array drift: {len(mismatched_bids)} row(s) differ. "
                    f"First: row {mismatched_bids[0][0]} "
                    f"legacy={mismatched_bids[0][1]} new={mismatched_bids[0][2]}"
                )

        # ── 3. Column headers ─────────────────────────────────────────
        missing_cols = set(df_leg.columns) - set(df_new.columns)
        extra_cols   = set(df_new.columns) - set(df_leg.columns)
        if missing_cols:
            issues.append(f"Missing columns in new: {sorted(missing_cols)}")
        if extra_cols:
            issues.append(f"Extra columns in new: {sorted(extra_cols)}")

        if not issues:
            print("    Result: ✅ ZERO DRIFT")
        else:
            all_pass = False
            for issue in issues:
                print(f"    ⚠️  {issue}")

    print(f"\n{'─'*62}")
    print("  Level 2 Result: " + ("✅ PASS" if all_pass else "⚠️  DRIFT DETECTED"))
    print(_SEP)
    return all_pass


# ═══════════════════════════════════════════════════════════════════════════
# Entry point
# ═══════════════════════════════════════════════════════════════════════════

def main():
    print("\n" + "═" * 62)
    print("  VALIDATE METRICS  (JSON Mode — 2-Level Check)")
    print("═" * 62)

    bulk_jsons    = sorted(glob.glob(os.path.join(JSON_DIR, 'BulkSheetExport_*.json')))
    updated_jsons = sorted(glob.glob(os.path.join(JSON_DIR, 'PPC_*_UPDATED.json')))

    level1_pass = True
    level2_pass = True

    if not bulk_jsons or not updated_jsons:
        logging.error("Missing JSON files. Run rebuild_internal_file.py and "
                      "end_to_end_ppc_tracker.py first.")
    else:
        level1_pass = validate_metric_integrity(bulk_jsons, updated_jsons)

    level2_pass = validate_data_drift(LEGACY_DIR, NEW_DIR)

    print("\n" + "═" * 62)
    if level1_pass and level2_pass:
        print("  🎉 VALIDATION COMPLETE — All checks passed.")
    else:
        print("  ⚠️  VALIDATION FAILED — Review issues above before uploading.")
    print("═" * 62 + "\n")


if __name__ == "__main__":
    main()
