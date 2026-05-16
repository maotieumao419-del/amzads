"""
auto_bulk_sync.py  —  Core Logic Layer
======================================
Role  : Reads campaign data from Amazon Bulk JSON, then syncs (adds/updates)
        keywords and placements into the PPC Internal JSON.

Input  : data/working_json/BulkSheetExport_*.json
         data/working_json/PPC_*.json
Output : data/working_json/PPC_*_synced.json
         data/final_xlsx/PPC_*.xlsx  (via excel_json_handler)

Zero Excel I/O inside this script — pure dict operations.
"""

import os
import sys
import json
import glob
import logging
import re

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, 'io_handlers'))
from excel_json_handler import json_to_internal_xlsx

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

JSON_DIR  = os.path.join(BASE_DIR, 'data', 'working_json')
FINAL_DIR = os.path.join(BASE_DIR, 'data', 'final_xlsx')
os.makedirs(FINAL_DIR, exist_ok=True)


# ─── Data helpers ────────────────────────────────────────────────────────────

def clean_sku(val: str) -> str:
    return str(val).strip('\r\n\t ') if val else ''


def _safe_str(val) -> str:
    s = str(val).strip() if val else ''
    return '' if s.lower() == 'nan' else s


def _match_type_from_note(note: str) -> str:
    """Extract match type from note string like '[[Exact]-[0.56]-[10T-0R-0P]]'."""
    import re as _re
    m = _re.search(r'\[\[(Exact|Phrase|Broad|targeting)\]', str(note), _re.IGNORECASE)
    return m.group(1).lower() if m else ''


# ─── Stage 1: Extract campaign blocks from Bulk JSON ─────────────────────────

def extract_campaign_blocks(records: list) -> dict:
    """
    Group flat bulk JSON rows by Campaign ID.
    Returns sku_map:
      {
        sku: {
          'Portfolio ID': str,
          'Portfolio Name': str,
          'Campaigns': {
            campaign_name: {
              'Keywords': [ {Text, MatchType, Bid, EntityType, NeedsReview} ],
              'Placements': {Top, PP, ROS}
            }
          }
        }
      }
    """
    from collections import defaultdict

    # Group by Campaign ID
    by_camp = defaultdict(list)
    for row in records:
        cid = _safe_str(row.get('Campaign ID') or row.get('Campaign Name', ''))
        by_camp[cid].append(row)

    sku_map = {}

    for camp_id, group in by_camp.items():
        # Campaign-level metadata
        camp_row = next((r for r in group if _safe_str(r.get('Entity', '')).lower() == 'campaign'), None)
        if not camp_row:
            camp_row = group[0]

        campaign_name = _safe_str(camp_row.get('Campaign Name', f'Unknown_{camp_id}'))
        portfolio_id  = _safe_str(camp_row.get('Portfolio ID', ''))
        portfolio_name = _safe_str(camp_row.get('Portfolio Name (Informational only)', ''))

        # Placements — stored as int strings to avoid '10.0' in notes
        placements = {'Top': '0', 'PP': '0', 'ROS': '0'}
        for row in group:
            if _safe_str(row.get('Entity', '')).lower() != 'bidding adjustment':
                continue
            p_type = _safe_str(row.get('Placement', '')).lower()
            if 'business' in p_type:
                continue
            raw_pct = row.get('Percentage', 0)
            try:
                pct = str(int(float(raw_pct or 0)))
            except (ValueError, TypeError):
                pct = '0'
            if 'top' in p_type:
                placements['Top'] = pct
            elif 'product page' in p_type or 'product_page' in p_type:
                placements['PP'] = pct
            elif 'rest of search' in p_type or 'rest_of_search' in p_type:
                placements['ROS'] = pct

        # SKUs from Product Ad rows
        skus = list({
            clean_sku(r.get('SKU', ''))
            for r in group
            if _safe_str(r.get('Entity', '')).lower() == 'product ad'
            and clean_sku(r.get('SKU', ''))
        })
        if not skus:
            fallback = campaign_name.split('_')[0].strip()
            skus = [fallback] if fallback else []
            needs_review = True
        else:
            needs_review = False

        # Keywords / Product Targets
        keywords = []
        seen = set()
        for row in group:
            entity = _safe_str(row.get('Entity', '')).lower()
            if entity == 'keyword':
                text = _safe_str(row.get('Keyword Text', ''))
                m_type = _safe_str(row.get('Match Type', ''))
                orig_ent = 'Keyword'
            elif entity == 'product targeting':
                text = _safe_str(row.get('Product Targeting Expression', ''))
                m_type = 'targeting'
                orig_ent = 'Product Targeting'
            else:
                continue

            if not text:
                continue

            key = (text, m_type.lower())
            if key in seen:
                continue
            seen.add(key)

            keywords.append({
                'Text':        text,
                'MatchType':   m_type,
                'Bid':         _safe_str(row.get('Bid', '')),
                'EntityType':  orig_ent,
                'NeedsReview': needs_review
            })

        # Cross-join SKUs × Keywords
        for sku in skus:
            if sku not in sku_map:
                sku_map[sku] = {
                    'Portfolio ID':   portfolio_id,
                    'Portfolio Name': portfolio_name,
                    'Campaigns': {}
                }
            if campaign_name not in sku_map[sku]['Campaigns']:
                sku_map[sku]['Campaigns'][campaign_name] = {
                    'Keywords':   [],
                    'Placements': placements
                }
            existing = {
                (kw['Text'], kw['MatchType'].lower())
                for kw in sku_map[sku]['Campaigns'][campaign_name]['Keywords']
            }
            for kw in keywords:
                if (kw['Text'], kw['MatchType'].lower()) not in existing:
                    sku_map[sku]['Campaigns'][campaign_name]['Keywords'].append(kw)
                    existing.add((kw['Text'], kw['MatchType'].lower()))

    return sku_map


# ─── Stage 2: Sync sku_map into PPC Internal JSON ────────────────────────────

def sync_internal_json(internal_data: dict, sku_map: dict) -> dict:
    """
    Merge sku_map data into the internal_data dict (in-memory, no xlsx).
    Returns the mutated internal_data.
    """
    listing_rows = internal_data.get('Listing', [])
    existing_skus = {row.get('SKU', '').strip() for row in listing_rows}

    # ── Update Listing ──────────────────────────────────────────────
    for sku, info in sku_map.items():
        if sku not in existing_skus:
            new_entry = {'SKU': sku, 'Portfolio ID': info['Portfolio ID']}
            listing_rows.append(new_entry)
            existing_skus.add(sku)
            logging.info(f"  [Listing] Added new SKU: {sku}")
    internal_data['Listing'] = listing_rows

    # ── Update / create SKU sheets ──────────────────────────────────
    for sku, info in sku_map.items():
        if not info['Campaigns']:
            continue

        tab = sku[:31]

        if tab not in internal_data:
            # Create brand-new SKU sheet structure
            internal_data[tab] = {
                'sku_title':    f'SKU: {sku}',
                'base_headers': ['STT', 'Campaign Name', 'Loại Campaign', 'Target', 'Ghi chú'],
                'date_blocks':  [],
                'rows':         []
            }

        sheet = internal_data[tab]
        if isinstance(sheet, list):
            # Tabular sheets (Listing, Portfolio ID) are not synced this way
            continue

        # Triple key: (Campaign, Target, MatchType) — each combination = 1 unique row
        existing_keys = {
            (
                r.get('Campaign Name', ''),
                r.get('Target', ''),
                _match_type_from_note(r.get('Ghi chu', '') or r.get('Ghi chú', ''))
            ): i
            for i, r in enumerate(sheet['rows'])
        }

        stt_counter = max(
            (r.get('STT', 0) for r in sheet['rows'] if isinstance(r.get('STT'), int)),
            default=0
        )

        for camp_name, details in info['Campaigns'].items():
            placements = details['Placements']
            note_base  = f"[{placements['Top']}T-{placements['ROS']}R-{placements['PP']}P]"

            for kw in details['Keywords']:
                text    = str(kw['Text']).strip()
                m_type  = str(kw['MatchType']).strip()
                ent     = kw.get('EntityType', 'Keyword')

                # Bid: keep as clean number string
                try:
                    bid = str(round(float(kw['Bid']), 2)) if kw['Bid'] else '0.0'
                except (ValueError, TypeError):
                    bid = '0.0'

                display_mt   = m_type if m_type and m_type.lower() not in ('nan', '') else 'N/A'
                note_content = f"[[{display_mt}]-[{bid}]-{note_base}]"

                # Triple key lookup
                key = (camp_name, text, display_mt.lower())
                if key in existing_keys:
                    # Update note (bid/placement may have changed)
                    idx = existing_keys[key]
                    sheet['rows'][idx]['Ghi chú'] = note_content
                    sheet['rows'][idx]['Status'] = 'Enable'
                else:
                    stt_counter += 1
                    col3 = 'Product Targeting' if ent == 'Product Targeting' else 'Keyword (Từ khóa)'
                    new_row = {
                        'STT':            stt_counter,
                        'Campaign Name':  camp_name,
                        'Loại Campaign':  col3,
                        'Target':         text,
                        'Ghi chú':        note_content,
                        'Status':         'Enable',
                        'metrics_by_date': {}
                    }
                    sheet['rows'].append(new_row)
                    existing_keys[key] = len(sheet['rows']) - 1

    return internal_data


# ─── Entry point ─────────────────────────────────────────────────────────────

def main():
    print("--- AUTO BULK SYNC (JSON Mode) ---")

    # 1. Load all Bulk JSONs
    bulk_jsons = sorted(glob.glob(os.path.join(JSON_DIR, 'BulkSheetExport_*.json')))
    if not bulk_jsons:
        logging.error("No BulkSheetExport_*.json found. Run rebuild_internal_file.py first.")
        return

    master_sku_map = {}
    for bj in bulk_jsons:
        with open(bj, 'r', encoding='utf-8') as f:
            records = json.load(f)
        file_map = extract_campaign_blocks(records)
        for sku, info in file_map.items():
            if sku not in master_sku_map:
                master_sku_map[sku] = info
            else:
                for camp, details in info['Campaigns'].items():
                    master_sku_map[sku]['Campaigns'][camp] = details

    logging.info(f"Extracted {len(master_sku_map)} SKUs from Bulk JSONs.")

    # 2. Load + sync each Internal JSON
    ppc_jsons = sorted(glob.glob(os.path.join(JSON_DIR, 'PPC_*.json')))
    ppc_jsons = [p for p in ppc_jsons if '_synced' not in p and '_UPDATED' not in p]

    if not ppc_jsons:
        logging.error("No PPC_*.json found. Run rebuild_internal_file.py first.")
        return

    for pj in ppc_jsons:
        basename = os.path.basename(pj)
        logging.info(f"Syncing: {basename}")

        with open(pj, 'r', encoding='utf-8') as f:
            internal_data = json.load(f)

        synced = sync_internal_json(internal_data, master_sku_map)

        # Save synced JSON
        synced_json_path = pj.replace('.json', '_synced.json')
        with open(synced_json_path, 'w', encoding='utf-8') as f:
            json.dump(synced, f, ensure_ascii=False, indent=2, default=str)
        logging.info(f"  → Saved synced JSON: {os.path.basename(synced_json_path)}")

        # Convert to xlsx in final_xlsx/
        xlsx_name    = basename.replace('.json', '.xlsx')
        xlsx_out     = os.path.join(FINAL_DIR, xlsx_name)
        ok = json_to_internal_xlsx(synced_json_path, xlsx_out)
        if ok:
            logging.info(f"  → Saved xlsx: {xlsx_name}")

    logging.info("SYNC COMPLETE.")


if __name__ == "__main__":
    main()
