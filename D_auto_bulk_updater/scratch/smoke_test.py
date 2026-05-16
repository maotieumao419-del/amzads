import sys
sys.path.insert(0, r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\io_handlers')
sys.path.insert(0, r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\core_logic')
import json, os

JSON_DIR = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\working_json'

# --- Run auto_bulk_sync ---
from auto_bulk_sync import extract_campaign_blocks, sync_internal_json
with open(os.path.join(JSON_DIR, 'BulkSheetExport_3004-0405.json'), encoding='utf-8') as f:
    records = json.load(f)
sku_map = extract_campaign_blocks(records)
print(f'extract_campaign_blocks: {len(sku_map)} SKUs')

with open(os.path.join(JSON_DIR, 'PPC_Musemory.json'), encoding='utf-8') as f:
    internal_data = json.load(f)
synced = sync_internal_json(internal_data, sku_map)
with open(os.path.join(JSON_DIR, 'PPC_Musemory_synced.json'), 'w', encoding='utf-8') as f:
    json.dump(synced, f, ensure_ascii=False, indent=2, default=str)
print('sync_internal_json: OK')

# --- Run end_to_end_ppc_tracker ---
from end_to_end_ppc_tracker import extract_full_metrics, inject_metrics
camp_map = extract_full_metrics(records)
print(f'extract_full_metrics: {len(camp_map)} campaigns')
updated = inject_metrics(synced, camp_map, '3004-0405')
with open(os.path.join(JSON_DIR, 'PPC_Musemory_UPDATED.json'), 'w', encoding='utf-8') as f:
    json.dump(updated, f, ensure_ascii=False, indent=2, default=str)
print('inject_metrics: OK')

# --- Check ONM_NURSE sheet metrics ---
nurse = updated.get('ONM_NURSE', {})
rows = nurse.get('rows', [])
print(f'ONM_NURSE rows: {len(rows)}')
for r in rows[:4]:
    m = r.get('metrics_by_date', {}).get('3004-0405', {})
    print(f"  Target={str(r.get('Target',''))[:35]:35} Imp={m.get('Impressions', 'N/A')}")

# --- Test json_to_internal_xlsx ---
from excel_json_handler import json_to_internal_xlsx
out_xlsx = os.path.join(r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\final_xlsx', 'PPC_Musemory_UPDATED.xlsx')
ok = json_to_internal_xlsx(os.path.join(JSON_DIR, 'PPC_Musemory_UPDATED.json'), out_xlsx)
print(f'json_to_internal_xlsx: {ok}')
