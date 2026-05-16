"""
end_to_end_ppc_tracker.py  —  Core Logic Layer
===============================================
Role  : Reads Bulk JSON metrics and injects them into the PPC Internal JSON
        at the correct (Campaign, Target, MatchType) row.

Input  : data/working_json/BulkSheetExport_*.json
         data/working_json/PPC_*_synced.json  (or PPC_*.json if no synced copy)
Output : data/working_json/PPC_*_UPDATED.json
         data/final_xlsx/PPC_*_UPDATED.xlsx  (via excel_json_handler)

Zero Excel I/O — pure dict operations.
"""

import os
import sys
import json
import glob
import logging
import re
from collections import defaultdict

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, 'io_handlers'))
from excel_json_handler import json_to_internal_xlsx

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

JSON_DIR    = os.path.join(BASE_DIR, 'data', 'working_json')
FINAL_DIR   = os.path.join(BASE_DIR, 'data', 'final_xlsx')
REPORTS_DIR = os.path.join(BASE_DIR, 'data', 'reports')
os.makedirs(FINAL_DIR,   exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

METRIC_COLUMNS = [
    'Impressions', 'Clicks', 'Click-through Rate', 'Spend', 'Sales',
    'Orders', 'Units', 'Conversion Rate', 'ACOS', 'CPC', 'ROAS'
]

_MATCH_RE = re.compile(r'\[\[(Exact|Phrase|Broad|targeting)\]', re.IGNORECASE)


def generate_unmatched_report(updated_json_path: str, bulk_json_path: str,
                              date_range: str, report_dir: str) -> str | None:
    """
    Generates an Excel report with two sheets:
      - Sheet "No-active" : rows in Internal that were NOT found in Bulk
      - Sheet "Bulk Only" : keywords in Bulk that were NOT matched to any Internal row
    Now with detailed reasons for Bulk-only items.
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter

        with open(updated_json_path, 'r', encoding='utf-8') as f:
            updated = json.load(f)
        with open(bulk_json_path, 'r', encoding='utf-8') as f:
            bulk_records = json.load(f)

        # ── 1. Map Internal Data for lookup ────────────────────────────────
        # Structure: internal_map[sku][campaign][(target, match)] = True
        internal_map = {}
        all_internal_campaigns = set() # To detect if campaign exists in another SKU
        campaign_to_sku_internal = {}  # To find which SKU a campaign belongs to in internal

        for sku, sheet_data in updated.items():
            if not isinstance(sheet_data, dict) or 'rows' not in sheet_data:
                continue
            internal_map[sku] = {}
            for row in sheet_data.get('rows', []):
                camp = _safe(row.get('Campaign Name'))
                target = _safe(row.get('Target')).lower()
                note = _safe(row.get('Ghi ch\u00fa') or row.get('Ghi chu') or '')
                m = _MATCH_RE.search(note)
                match = m.group(1).capitalize() if m else ''
                
                if camp:
                    all_internal_campaigns.add(camp)
                    if camp not in internal_map[sku]:
                        internal_map[sku][camp] = set()
                    internal_map[sku][camp].add((target, match))
                    campaign_to_sku_internal[camp] = sku

        # ── 2. Build SKU map from Bulk (via Product Ads) ──────────────────
        bulk_campaign_to_sku = {}
        for rec in bulk_records:
            entity = _safe(rec.get('Entity')).lower()
            if entity == 'product ad':
                camp = _safe(rec.get('Campaign Name'))
                sku = _safe(rec.get('SKU'))
                if camp and sku:
                    bulk_campaign_to_sku[camp] = sku

        # ── 3. Collect No-active rows (Internal missing from Bulk) ──────────
        bulk_keys = set()
        for rec in bulk_records:
            entity = _safe(rec.get('Entity')).lower()
            if entity not in ('keyword', 'product targeting'):
                continue
            camp  = _safe(rec.get('Campaign Name'))
            text  = _safe(rec.get('Keyword Text') or rec.get('Product Targeting Expression')).lower()
            match = _safe(rec.get('Match Type')).capitalize()
            if camp and text:
                bulk_keys.add((camp, text, match))

        no_active_rows = []
        for sku, sheet_data in updated.items():
            if not isinstance(sheet_data, dict) or 'rows' not in sheet_data:
                continue
            for row in sheet_data.get('rows', []):
                if row.get('Status') != 'No-active':
                    continue
                camp   = _safe(row.get('Campaign Name'))
                target = _safe(row.get('Target')).lower()
                note   = _safe(row.get('Ghi ch\u00fa') or row.get('Ghi chu') or '')
                m      = _MATCH_RE.search(note)
                match  = m.group(1).capitalize() if m else ''

                no_active_rows.append({
                    'SKU':          sku,
                    'Campaign Name': camp,
                    'Loại':         _safe(row.get('Loại Campaign')),
                    'Target':       target,
                    'Match Type':   match,
                    'Ghi chú':      note,
                    'Reason':       'Not found in Bulk Export',
                })

        # ── 4. Collect Bulk-only rows with detailed reasons ────────────────
        bulk_only_rows = []
        for rec in bulk_records:
            entity = _safe(rec.get('Entity')).lower()
            if entity not in ('keyword', 'product targeting'):
                continue
            camp  = _safe(rec.get('Campaign Name'))
            text  = _safe(rec.get('Keyword Text') or rec.get('Product Targeting Expression'))
            match = _safe(rec.get('Match Type')).capitalize()
            impr  = rec.get('Impressions', 0)
            
            if not camp:
                bulk_only_rows.append({
                    'Campaign Name': 'EMPTY',
                    'Target': text,
                    'Match Type': match,
                    'Impressions': impr,
                    'Clicks': rec.get('Clicks', 0),
                    'Spend': rec.get('Spend', 0),
                    'Predicted SKU': 'N/A',
                    'Reason': 'Campaign Name bị trống trong file Bulk'
                })
                continue

            # Check if (camp, target, match) exists in internal
            found = False
            # Check all sheets because a campaign might be incorrectly placed
            for sku_sheet in internal_map:
                if camp in internal_map[sku_sheet]:
                    if (text.lower(), match) in internal_map[sku_sheet][camp]:
                        found = True
                        break
            
            if not found:
                # Analyze why it's missing
                reason = "Unknown"
                predicted_sku = bulk_campaign_to_sku.get(camp) or camp.split('_')[0].strip()
                tab_name = predicted_sku[:31]

                if tab_name not in internal_map:
                    reason = f"Sheet SKU '{tab_name}' chưa tồn tại trong file Internal"
                elif camp not in internal_map[tab_name]:
                    if camp in all_internal_campaigns:
                        actual_sku = campaign_to_sku_internal.get(camp)
                        reason = f"Campaign đang nằm sai SKU (Trong Internal: '{actual_sku}', Amazon: '{tab_name}')"
                    else:
                        reason = f"Campaign chưa có trong Sheet SKU '{tab_name}'"
                else:
                    reason = f"Thiếu Target/Match Type trong Campaign '{camp}'"

                bulk_only_rows.append({
                    'Campaign Name':  camp,
                    'Target':         text,
                    'Match Type':     match,
                    'Impressions':    impr,
                    'Clicks':         rec.get('Clicks', 0),
                    'Spend':          rec.get('Spend', 0),
                    'Predicted SKU':  tab_name,
                    'Reason':         reason,
                })

        # ── 5. Write Excel report ──────────────────────────────────────────
        wb = openpyxl.Workbook()
        HEADER_FILL = PatternFill('solid', fgColor='1F3864')
        HEADER_FONT = Font(bold=True, color='FFFFFF')
        WARN_FILL   = PatternFill('solid', fgColor='FFF2CC')

        def write_sheet(ws, headers, rows_data, color_col=None):
            for c_idx, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=c_idx, value=h)
                cell.font   = HEADER_FONT
                cell.fill   = HEADER_FILL
                cell.alignment = Alignment(horizontal='center', wrap_text=True)
            for r_idx, rd in enumerate(rows_data, 2):
                for c_idx, h in enumerate(headers, 1):
                    val = rd.get(h, '')
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    if color_col and h == color_col:
                        cell.fill = WARN_FILL
            for col_cells in ws.columns:
                col_letter = get_column_letter(col_cells[0].column)
                max_len = max((len(str(c.value or '')) for c in col_cells), default=8)
                ws.column_dimensions[col_letter].width = min(max_len + 4, 60)
            ws.row_dimensions[1].height = 22

        # Sheet 1: No-active
        ws1 = wb.active
        ws1.title = 'No-active (Missing from Bulk)'
        h1 = ['SKU', 'Campaign Name', 'Loại', 'Target', 'Match Type', 'Ghi chú', 'Reason']
        write_sheet(ws1, h1, no_active_rows, color_col='Reason')

        # Sheet 2: Bulk Only
        ws2 = wb.create_sheet('Bulk-Only (Not in Internal)')
        h2 = ['Campaign Name', 'Target', 'Match Type', 'Impressions', 'Clicks', 'Spend', 'Predicted SKU', 'Reason']
        write_sheet(ws2, h2, bulk_only_rows, color_col='Reason')

        # Save
        base = os.path.splitext(os.path.basename(updated_json_path))[0]
        report_path = os.path.join(report_dir, f"{base}_UNMATCHED_{date_range}.xlsx")
        wb.save(report_path)
        logging.info(f"  → Unmatched report updated: {os.path.basename(report_path)}")
        return report_path

    except Exception as e:
        logging.error(f"[generate_unmatched_report] Failed: {e}")
        import traceback; traceback.print_exc()
        return None



def _safe(val) -> str:
    s = str(val).strip() if val else ''
    return '' if s.lower() == 'nan' else s


# ─── Stage 1: Extract full metrics from Bulk JSON ────────────────────────────

def extract_full_metrics(records: list) -> dict:
    """
    Build campaign_map:
      {
        campaign_name: {
          'SKUs':       [sku, ...],
          'Placements': {Top, PP, ROS},
          'Keywords': [
            {Text, Match, EntityType, Impressions, Clicks, ...}
          ]
        }
      }
    """
    by_camp = defaultdict(list)
    for row in records:
        cid = _safe(row.get('Campaign ID') or row.get('Campaign Name', ''))
        by_camp[cid].append(row)

    campaign_map = {}

    for camp_id, group in by_camp.items():
        # Campaign name
        camp_row = next(
            (r for r in group if _safe(r.get('Entity', '')).lower() == 'campaign'),
            group[0]
        )
        campaign_name = _safe(camp_row.get('Campaign Name', f'Unknown_{camp_id}'))

        # Placements
        placements = {'Top': '0', 'PP': '0', 'ROS': '0'}
        for row in group:
            if _safe(row.get('Entity', '')).lower() != 'bidding adjustment':
                continue
            p_type = _safe(row.get('Placement', '')).lower()
            pct    = _safe(row.get('Percentage', '0')).replace('%', '') or '0'
            if 'top' in p_type:
                placements['Top'] = pct
            elif 'product page' in p_type or 'product_page' in p_type:
                placements['PP'] = pct
            elif 'rest of search' in p_type or 'rest_of_search' in p_type:
                placements['ROS'] = pct

        # SKUs
        skus = list({
            _safe(r.get('SKU', ''))
            for r in group
            if _safe(r.get('Entity', '')).lower() == 'product ad'
            and _safe(r.get('SKU', ''))
        })
        if not skus:
            skus = [campaign_name.split('_')[0].strip()]

        # Keywords / targets with metrics
        keywords = []
        for row in group:
            entity = _safe(row.get('Entity', '')).lower()
            if entity == 'keyword':
                text     = _safe(row.get('Keyword Text', ''))
                match    = _safe(row.get('Match Type', ''))
                ent_type = 'Keyword'
            elif entity == 'product targeting':
                text     = _safe(row.get('Product Targeting Expression', ''))
                match    = ''
                ent_type = 'Product Targeting'
            else:
                continue
            if not text:
                continue

            kw_data = {'Text': text, 'Match': match, 'EntityType': ent_type}
            for col in METRIC_COLUMNS:
                raw = row.get(col, 0)
                try:
                    kw_data[col] = float(raw) if raw not in (None, '', 'nan') else 0.0
                except (ValueError, TypeError):
                    kw_data[col] = 0.0
            keywords.append(kw_data)

        campaign_map[campaign_name] = {
            'SKUs':       skus,
            'Placements': placements,
            'Keywords':   keywords
        }

    return campaign_map


# ─── Stage 2: Inject metrics into Internal JSON ───────────────────────────────

def inject_metrics(internal_data: dict, campaign_map: dict, date_range: str) -> dict:
    """
    For each campaign in campaign_map, locate matching rows in the relevant
    SKU sheets of internal_data using key (Campaign Name, Target Text, Match Type)
    and write the metrics into metrics_by_date[date_range].
    """
    # Pre-build row index for every SKU sheet:
    #   sheet_index[sheet_name][(camp, target, match)] = row_list_index
    sheet_index = {}
    for sheet_name, sheet_data in internal_data.items():
        if not isinstance(sheet_data, dict) or 'rows' not in sheet_data:
            continue
        index = {}
        for i, row in enumerate(sheet_data.get('rows', [])):
            row['Status'] = 'No-active'  # Default if not found in master
            row['metrics_by_date'] = {}  # Clear stats as requested ("để trống")
            camp   = str(row.get('Campaign Name', '')).strip()
            target = str(row.get('Target', '')).strip()
            # Support unicode 'Ghi chú' key stored in JSON
            note   = str(row.get('Ghi ch\u00fa') or row.get('Ghi chu') or '')
            m      = _MATCH_RE.search(note)
            match  = m.group(1).capitalize() if m else ''
            if camp and target:
                # Triple key: same target with different match type = different row
                index[(camp, target, match)] = i
        sheet_index[sheet_name] = index

    for c_name, details in campaign_map.items():
        for sku in details['SKUs']:
            tab = sku[:31]
            if tab not in internal_data or tab not in sheet_index:
                continue

            sheet = internal_data[tab]

            # Ensure this date_range is registered in date_blocks
            if date_range not in sheet.get('date_blocks', []):
                sheet.setdefault('date_blocks', []).append(date_range)

            for kw in details['Keywords']:
                kw_match = kw['Match'].capitalize() \
                    if kw['Match'].lower() in ('exact', 'phrase', 'broad') else ''
                key = (c_name, kw['Text'], kw_match)

                if key not in sheet_index[tab]:
                    continue

                row_idx = sheet_index[tab][key]
                row     = sheet['rows'][row_idx]

                row.setdefault('metrics_by_date', {})[date_range] = {
                    m: kw.get(m, 0) for m in METRIC_COLUMNS
                }
                row['metrics_by_date'][date_range]['Placement Note'] = ''
                row['Status'] = 'Enable'

    # ─── Final Sorting ──────────────────────────────────────────────────
    for sheet_name, sheet_data in internal_data.items():
        if not isinstance(sheet_data, dict) or 'rows' not in sheet_data:
            continue
        
        def sort_rank(row):
            # 1. Status: Enable (0) < No-active (1)
            s_rank = 0 if row.get('Status') == 'Enable' else 1

            # 2. Entity type: Keyword (0) before Product Targeting (1)
            loai = str(row.get('Lo\u1ea1i Campaign') or '').strip().lower()
            e_rank = 1 if 'product targeting' in loai else 0

            # 3. Keyword/Target name (normalized for grouping)
            target = str(row.get('Target', '')).strip().casefold()

            # 4. Match Type Rank: Exact(0) > Phrase(1) > Broad(2) > Targeting(3) > Others(4)
            note = str(row.get('Ghi ch\u00fa') or row.get('Ghi chu') or '')
            m = _MATCH_RE.search(note)
            mt = m.group(1).lower() if m else ''

            mt_rank = 4
            if mt == 'exact':        mt_rank = 0
            elif mt == 'phrase':     mt_rank = 1
            elif mt == 'broad':      mt_rank = 2
            elif mt == 'targeting':  mt_rank = 3

            return (s_rank, e_rank, target, mt_rank)

        sheet_data['rows'].sort(key=sort_rank)

        # 4. Re-number STT to keep it clean after sort
        for i, row in enumerate(sheet_data['rows'], 1):
            row['STT'] = i

    return internal_data


# ─── Entry point ─────────────────────────────────────────────────────────────

def _get_date_range(filename: str) -> str:
    import re as _re
    m = _re.search(r'(\d{4}-\d{4})', filename)
    return m.group(1) if m else 'unknown'


def main():
    print("--- END-TO-END PPC TRACKER (JSON Mode) ---")

    # 1. Load Bulk JSONs and extract metrics
    bulk_jsons = sorted(glob.glob(os.path.join(JSON_DIR, 'BulkSheetExport_*.json')))
    if not bulk_jsons:
        logging.error("No BulkSheetExport_*.json found. Run rebuild_internal_file.py first.")
        return

    # 2. Process each Internal JSON
    # Prefer _synced versions; fall back to base PPC_.json
    all_ppc = glob.glob(os.path.join(JSON_DIR, 'PPC_*.json'))
    synced  = [p for p in all_ppc if p.endswith('_synced.json')]
    base    = [p for p in all_ppc if not p.endswith('_synced.json')
               and '_UPDATED' not in p]
    ppc_jsons = synced if synced else base

    if not ppc_jsons:
        logging.error("No PPC JSON found. Run auto_bulk_sync.py first.")
        return

    for bj in bulk_jsons:
        date_range = _get_date_range(os.path.basename(bj))
        with open(bj, 'r', encoding='utf-8') as f:
            records = json.load(f)
        campaign_map = extract_full_metrics(records)
        logging.info(f"Bulk {os.path.basename(bj)}: {len(campaign_map)} campaigns, date={date_range}")

        for pj in ppc_jsons:
            basename = os.path.basename(pj)
            logging.info(f"  Injecting → {basename}")

            with open(pj, 'r', encoding='utf-8') as f:
                internal_data = json.load(f)

            updated = inject_metrics(internal_data, campaign_map, date_range)

            # Save UPDATED JSON
            base_name  = basename.replace('_synced', '').replace('.json', '')
            out_json   = os.path.join(JSON_DIR, f"{base_name}_UPDATED.json")
            with open(out_json, 'w', encoding='utf-8') as f:
                json.dump(updated, f, ensure_ascii=False, indent=2, default=str)
            logging.info(f"  → JSON: {os.path.basename(out_json)}")

            # Convert to xlsx
            out_xlsx = os.path.join(FINAL_DIR, f"{base_name}_UPDATED.xlsx")
            ok = json_to_internal_xlsx(out_json, out_xlsx)
            if ok:
                logging.info(f"  → xlsx: {os.path.basename(out_xlsx)}")

            # Generate Unmatched Report
            generate_unmatched_report(out_json, bj, date_range, REPORTS_DIR)

    logging.info("TRACKER COMPLETE.")


if __name__ == "__main__":
    main()
