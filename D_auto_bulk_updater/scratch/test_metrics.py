import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'core_logic'))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'io_handlers'))

from auto_bulk_updater import clean_metric, extract_full_metrics

# --- clean_metric ---
assert clean_metric('$1,234.56')   == 1234.56, 'FAIL: dollar sign'
assert clean_metric('12.5%', float)== 12.5,    'FAIL: percent'
assert clean_metric('1,200', int)  == 1200,    'FAIL: comma int'
assert clean_metric(None)          == 0.0,     'FAIL: None'
assert clean_metric('--')          == 0.0,     'FAIL: dash'
assert clean_metric(0)             == 0.0,     'FAIL: zero int'
print('clean_metric: ALL PASS')

# --- extract_full_metrics: interpolation (all rates missing) ---
raw = {
    'Impressions': 1000,
    'Clicks': 20,
    'Click-through Rate': 0.0,
    'Spend': '5.00',
    'Sales': '25.00',
    'Orders': 2,
    'Units': 2,
    'Conversion Rate': 0.0,
    'ACOS': 0.0,
    'CPC': 0.0,
    'ROAS': 0.0,
    'Placement Note': ''
}
fm = extract_full_metrics(raw)
print('full_metrics:', fm)

assert fm['impressions'] == 1000
assert fm['clicks']      == 20
ctr_ok = abs(fm['ctr']  - 2.0)  < 0.001
assert ctr_ok,  'CTR interpolation failed: ' + str(fm['ctr'])
assert fm['spend']  == 5.0
assert fm['sales']  == 25.0
assert fm['orders'] == 2
assert fm['units']  == 2
cvr_ok = abs(fm['cvr']  - 10.0) < 0.001
assert cvr_ok,  'CVR interpolation failed: ' + str(fm['cvr'])
acos_ok= abs(fm['acos'] - 20.0) < 0.001
assert acos_ok, 'ACOS interpolation failed: ' + str(fm['acos'])
cpc_ok = abs(fm['cpc']  - 0.25) < 0.001
assert cpc_ok,  'CPC interpolation failed: '  + str(fm['cpc'])
roas_ok= abs(fm['roas'] - 5.0)  < 0.001
assert roas_ok, 'ROAS interpolation failed: ' + str(fm['roas'])
print('extract_full_metrics interpolation: ALL PASS')

# --- Non-zero rates must NOT be overwritten ---
raw2 = {
    'Impressions': 1000, 'Clicks': 20, 'Click-through Rate': 1.5,
    'Spend': 5, 'Sales': 25, 'Orders': 2, 'Units': 2,
    'Conversion Rate': 8.0, 'ACOS': 19.0, 'CPC': 0.22, 'ROAS': 4.8
}
fm2 = extract_full_metrics(raw2)
assert fm2['ctr']  == 1.5,  'CTR should use provided value'
assert fm2['cvr']  == 8.0,  'CVR should use provided value'
assert fm2['acos'] == 19.0, 'ACOS should use provided value'
assert fm2['cpc']  == 0.22, 'CPC should use provided value'
assert fm2['roas'] == 4.8,  'ROAS should use provided value'
print('Non-zero rates not overwritten: ALL PASS')

print()
print('=== ALL UNIT TESTS PASSED ===')
