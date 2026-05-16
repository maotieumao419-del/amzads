import json
from collections import Counter

with open(r'data\working_json\Amazon_Upload_Ready_15052026.json', encoding='utf-8') as f:
    rows = json.load(f)

entities = Counter(r.get('Entity', '') for r in rows)
states   = Counter(r.get('State',  '') for r in rows)

print(f'Total rows: {len(rows)}')
print()
print('Entity breakdown:')
for k, v in entities.most_common():
    print(f'  {k:<26} {v}')
print()
print('State breakdown:')
for k, v in states.most_common():
    s = repr(k)
    print(f'  {s:<15} {v}')
print()

# Bid-cap rows = Keyword entity + has Bid value + no State change
bid_rows = [
    r for r in rows
    if r.get('Bid') and r.get('Entity') == 'Keyword' and r.get('State', '') == ''
]
print(f'Keyword rows with Bid Capped (set_bid): {len(bid_rows)}')
print('Sample (first 3):')
for r in bid_rows[:3]:
    camp = r.get('Campaign Name', '')[:40]
    kw   = r.get('Keyword Text', '')
    bid  = r.get('Bid', '')
    print(f'  Camp={camp!r}  KW={kw!r}  Bid={bid!r}')

# Integrity: BA rows must have blank KW/MT/Bid
ba = [r for r in rows if r.get('Entity', '') == 'Bidding Adjustment']
ba_dirty = [r for r in ba if r.get('Keyword Text') or r.get('Match Type') or r.get('Bid')]
print()
print(f'BA rows: {len(ba)} — dirty (KW/MT/Bid non-blank): {len(ba_dirty)}  <-- must be 0')
