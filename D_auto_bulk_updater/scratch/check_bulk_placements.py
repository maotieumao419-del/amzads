
import json
import os
from collections import defaultdict

BASE_DIR = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater'
JSON_DIR = os.path.join(BASE_DIR, 'data', 'working_json')
BULK_FILE = os.path.join(JSON_DIR, 'BulkSheetExport_3004-0405.json')

def check_placements():
    with open(BULK_FILE, 'r', encoding='utf-8') as f:
        bulk = json.load(f)

    camps_with_pct = []
    total_camps = set()

    by_camp = defaultdict(list)
    for row in bulk:
        camp = row.get('Campaign Name')
        if camp:
            total_camps.add(camp)
            by_camp[camp].append(row)

    for camp, group in by_camp.items():
        placements = {'Top': 0, 'PP': 0, 'ROS': 0}
        has_adj = False
        for row in group:
            entity = str(row.get('Entity', '')).lower()
            if entity == 'bidding adjustment':
                has_adj = True
                p_type = str(row.get('Placement', '')).lower()
                pct = row.get('Percentage', 0)
                try:
                    pct_val = float(pct)
                except:
                    pct_val = 0
                
                if 'top' in p_type: placements['Top'] = pct_val
                elif 'product page' in p_type or 'product_page' in p_type: placements['PP'] = pct_val
                elif 'rest of search' in p_type or 'rest_of_search' in p_type: placements['ROS'] = pct_val

        if any(v > 0 for v in placements.values()):
            camps_with_pct.append({
                'name': camp,
                'pct': placements
            })

    print(f"Total campaigns in Bulk: {len(total_camps)}")
    print(f"Campaigns with non-zero placement adjustments: {len(camps_with_pct)}")
    
    if camps_with_pct:
        print("\nSample campaigns with adjustments:")
        for item in camps_with_pct[:10]:
            print(f" - {item['name']}: {item['pct']}")
    else:
        print("\nWARNING: No non-zero placement adjustments found in Bulk JSON!")

if __name__ == '__main__':
    check_placements()
