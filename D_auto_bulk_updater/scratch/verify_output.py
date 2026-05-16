import json

path = r'data\working_json\Amazon_Upload_Ready_15052026.json'
with open(path, 'r', encoding='utf-8') as f:
    rows = json.load(f)

print(f'Total rows: {len(rows)}')
print()

ba_rows = [r for r in rows if str(r.get('Entity', '')).lower() == 'bidding adjustment']
kw_rows = [r for r in rows if str(r.get('Entity', '')).lower() == 'keyword']
pt_rows = [r for r in rows if str(r.get('Entity', '')).lower() == 'product targeting']

print('Entity breakdown:')
print(f'  Bidding Adjustment : {len(ba_rows)}')
print(f'  Keyword            : {len(kw_rows)}')
print(f'  Product Targeting  : {len(pt_rows)}')
print()

VERIFY_COLS = ['Entity', 'Operation', 'Campaign Name', 'Placement', 'Percentage',
               'Keyword Text', 'Match Type', 'Bid', 'State']

print('=== Bidding Adjustment rows (first 3) ===')
for i, r in enumerate(ba_rows[:3], 1):
    print(f'Row {i}:')
    for col in VERIFY_COLS:
        val = r.get(col, '')
        print(f'  {col:<22} = {repr(val)}')
    print()

print('=== Integrity Checks ===')
ba_fail_op    = [r for r in ba_rows if r.get('Operation') != 'Update']
ba_fail_plc   = [r for r in ba_rows if not r.get('Placement')]
ba_fail_pct   = [r for r in ba_rows if r.get('Percentage', '') == '']
ba_fail_blank = [r for r in ba_rows if r.get('Keyword Text') or r.get('Match Type') or r.get('Bid')]

print(f'  BA rows where Operation != "Update"    : {len(ba_fail_op)}')
print(f'  BA rows with empty Placement           : {len(ba_fail_plc)}')
print(f'  BA rows with empty Percentage          : {len(ba_fail_pct)}')
print(f'  BA rows with non-blank KW / MT / Bid   : {len(ba_fail_blank)}   <-- must be 0')
print()

print('=== Paused Keyword rows (first 1) ===')
paused = [r for r in kw_rows if str(r.get('State', '')).lower() == 'paused']
print(f'  Total Paused Keyword rows: {len(paused)}')
if paused:
    for col in VERIFY_COLS:
        empty = ''
        print(f'  {col:<22} = {repr(paused[0].get(col, empty))}')
