
import json
import os
import re

BASE_DIR = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater'
JSON_DIR = os.path.join(BASE_DIR, 'data', 'working_json')

BULK_FILE = os.path.join(JSON_DIR, 'BulkSheetExport_3004-0405.json')
INTERNAL_FILE = os.path.join(JSON_DIR, 'PPC_Musemory_UPDATED.json')

def _safe(val):
    s = str(val).strip() if val else ''
    return '' if s.lower() == 'nan' else s

def analyze():
    with open(INTERNAL_FILE, 'r', encoding='utf-8') as f:
        internal = json.load(f)
    with open(BULK_FILE, 'r', encoding='utf-8') as f:
        bulk = json.load(f)

    # 1. Map internal data for fast lookup
    # Structure: internal_map[sku][campaign][target][match] = True
    internal_map = {}
    sku_list = internal.keys()
    
    # We also need a flat list of all campaigns in internal
    internal_all_camps = set()
    internal_sku_camps = {} # sku -> set of camps

    match_re = re.compile(r'\[\[(Exact|Phrase|Broad|targeting)\]', re.IGNORECASE)

    for sku, data in internal.items():
        if not isinstance(data, dict) or 'rows' not in data: continue
        internal_map[sku] = {}
        internal_sku_camps[sku] = set()
        for row in data['rows']:
            camp = _safe(row.get('Campaign Name'))
            target = _safe(row.get('Target')).lower()
            note = _safe(row.get('Ghi chú') or row.get('Ghi chu') or '')
            m = match_re.search(note)
            match = m.group(1).capitalize() if m else ''
            
            if camp:
                internal_all_camps.add(camp)
                internal_sku_camps[sku].add(camp)
                if camp not in internal_map[sku]:
                    internal_map[sku][camp] = {}
                internal_map[sku][camp][(target, match)] = True

    # 2. Analyze Bulk records
    # Group by campaign first to find SKU
    by_camp = {}
    for row in bulk:
        c_name = _safe(row.get('Campaign Name'))
        if not c_name: continue
        if c_name not in by_camp: by_camp[c_name] = []
        by_camp[c_name].append(row)

    bulk_only = []

    for c_name, group in by_camp.items():
        # Find SKUs for this campaign
        skus = list({
            _safe(r.get('SKU')) for r in group 
            if _safe(r.get('Entity')).lower() == 'product ad' and _safe(r.get('SKU'))
        })
        if not skus:
            skus = [c_name.split('_')[0].strip()]
        
        for row in group:
            entity = _safe(row.get('Entity')).lower()
            if entity not in ('keyword', 'product targeting'): continue
            
            target = _safe(row.get('Keyword Text') or row.get('Product Targeting Expression')).lower()
            match = _safe(row.get('Match Type')).strip().capitalize()
            
            # Check if this (camp, target, match) is in internal for ANY of the possible SKUs
            found = False
            reasons = []
            
            for sku in skus:
                tab = sku[:31]
                if tab in internal_map:
                    if c_name in internal_map[tab]:
                        if (target, match) in internal_map[tab][c_name]:
                            found = True
                            break
                        else:
                            reasons.append(f"Target '{target}' [{match}] missing in SKU '{tab}' / Campaign '{c_name}'")
                    else:
                        reasons.append(f"Campaign '{c_name}' missing in SKU sheet '{tab}'")
                else:
                    reasons.append(f"SKU sheet '{tab}' missing in Internal file")
            
            if not found:
                # Deduplicate and pick the most relevant reason
                # Priority: SKU missing > Campaign missing > Target missing
                final_reason = "Unknown"
                if any("SKU sheet" in r for r in reasons):
                    final_reason = "Thiếu Sheet SKU trong file Internal"
                elif any("Campaign" in r for r in reasons):
                    final_reason = "Thiếu Campaign trong Sheet SKU"
                else:
                    final_reason = "Thiếu Target/Match Type trong Campaign"

                bulk_only.append({
                    'Campaign': c_name,
                    'Target': target,
                    'Match': match,
                    'SKUs': skus,
                    'Reason': final_reason
                })

    # Summary
    print(f"Total Bulk-only items found: {len(bulk_only)}")
    reason_counts = {}
    for item in bulk_only:
        reason_counts[item['Reason']] = reason_counts.get(item['Reason'], 0) + 1
    
    print("\nBreakdown by Reason:")
    for r, count in reason_counts.items():
        print(f" - {r}: {count}")

    # Sample items
    if bulk_only:
        print("\nSample items:")
        for item in bulk_only[:10]:
            print(f"  [{item['Reason']}] {item['Campaign']} | {item['Target']} | {item['Match']}")

if __name__ == '__main__':
    analyze()
