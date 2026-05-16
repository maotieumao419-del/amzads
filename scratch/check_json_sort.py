import json
import os

path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\working_json\PPC_Musemory_UPDATED.json'
with open(path, 'r', encoding='utf-8') as f:
    d = json.load(f)

s = d['MOTHERDAY_ILOVEYOUMOM']
print(f"Sheet: MOTHERDAY_ILOVEYOUMOM")
print(f"Rows count: {len(s['rows'])}")
for i, r in enumerate(s['rows'][:10]):
    print(f"  {i}: Target='{r.get('Target')}', Status='{r.get('Status')}', Note='{r.get('Ghi chú')}'")
