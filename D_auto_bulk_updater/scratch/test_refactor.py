"""Smoke test for the refactored auto_bulk_updater.py (3 modules)."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'core_logic'))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'io_handlers'))

from auto_bulk_updater import (
    clean_metric, extract_full_metrics,
    _static_evaluate, run_rule_engine, build_upload_rows
)

# ─── Module 1: clean_metric ─────────────────────────────────────────────────
assert clean_metric('$1,234.56')    == 1234.56
assert clean_metric('12.5%', float) == 12.5
assert clean_metric('1,200', int)   == 1200
assert clean_metric(None)           == 0.0
assert clean_metric('--')           == 0.0
print("M1 clean_metric: PASS")

# ─── Module 1: extract_full_metrics with interpolation ──────────────────────
raw = {
    'Impressions': 1000, 'Clicks': 20,
    'Click-through Rate': 0.0,   # → 20/1000*100 = 2.0
    'Spend': '5.00', 'Sales': '25.00',
    'Orders': 2, 'Units': 2,
    'Conversion Rate': 0.0,      # → 2/20*100 = 10.0
    'ACOS': 0.0,                 # → 5/25*100 = 20.0
    'CPC': 0.0,                  # → 5/20 = 0.25
    'ROAS': 0.0,                 # → 25/5 = 5.0
}
fm = extract_full_metrics(raw)
assert fm['impressions'] == 1000
assert fm['clicks']      == 20
assert abs(fm['ctr']  - 2.0)  < 0.001, fm['ctr']
assert fm['spend']  == 5.0
assert fm['sales']  == 25.0
assert fm['orders'] == 2
assert abs(fm['cvr']  - 10.0) < 0.001, fm['cvr']
assert abs(fm['acos'] - 20.0) < 0.001, fm['acos']
assert abs(fm['cpc']  - 0.25) < 0.001, fm['cpc']
assert abs(fm['roas'] - 5.0)  < 0.001, fm['roas']
print("M1 extract_full_metrics interpolation: PASS")

# Provided non-zero rates must NOT be overwritten
raw2 = dict(raw, **{'Click-through Rate':1.5,'Conversion Rate':8.0,
                     'ACOS':19.0,'CPC':0.22,'ROAS':4.8})
fm2 = extract_full_metrics(raw2)
assert fm2['ctr'] == 1.5 and fm2['cvr'] == 8.0
assert fm2['acos'] == 19.0 and fm2['cpc'] == 0.22 and fm2['roas'] == 4.8
print("M1 non-zero rates preserved: PASS")

# ─── Module 2: _static_evaluate — Bidding Adjustment via match_type_logic ───
rule_ba = {
    "rule_name": "Optimize_Placement_No_Order",
    "conditions": {"impressions_min": 1000, "clicks_min": 10, "orders_max": 0},
    "match_type_logic": {
        "Broad": {"target_entity": "Bidding Adjustment",
                  "placement_type": "Placement Product Page",
                  "set_percentage": 0, "action": "Update"},
        "Phrase": {"target_entity": "Bidding Adjustment",
                   "placement_type": "Placement Product Page",
                   "set_percentage": 10, "action": "Update"}
    },
    "log_template": "Placement Adjusted: {match_type}"
}
fm_ba = extract_full_metrics({'Impressions':1500,'Clicks':15,
                               'Orders':0,'Units':0,'Spend':3,'Sales':0,
                               'Click-through Rate':1,'Conversion Rate':0,
                               'ACOS':0,'CPC':0,'ROAS':0})
results_ba = _static_evaluate(rule_ba, fm_ba, 'Broad')
assert len(results_ba) == 1
assert results_ba[0]['target_entity'] == 'Bidding Adjustment'
assert results_ba[0]['placement_type'] == 'Placement Product Page'
assert results_ba[0]['set_percentage'] == 0
print("M2 static_evaluate Bidding Adjustment: PASS")

# _static_evaluate — flat Paused rule
rule_flat = {
    "rule_name": "Fatal_Bleeder_Zero_Sales",
    "conditions": {"clicks_min": 15, "orders_max": 0},
    "action": "Paused",
    "target_entity": "Keyword",
    "log_template": "Paused: {clicks} Clicks"
}
results_flat = _static_evaluate(rule_flat, fm_ba, 'Broad')
assert len(results_flat) == 1
assert results_flat[0]['action'] == 'Paused'
assert results_flat[0]['target_entity'] == 'Keyword'
print("M2 static_evaluate flat Keyword: PASS")

# run_rule_engine — no script files exist, should use static path silently
rules_cfg = {"rules": [rule_ba, rule_flat]}
actions = run_rule_engine(fm_ba, 'Broad', rules_cfg)
# Should get 2 results: 1 Bidding Adjustment + 1 Paused Keyword
assert len(actions) == 2, f"Expected 2, got {len(actions)}"
print("M2 run_rule_engine (static fallback): PASS")

# ─── Module 3: build_upload_rows ────────────────────────────────────────────
triggered = [
    {
        'SKU': 'TEST_SKU', 'Campaign Name': 'Camp A',
        'Target': 'nurse gifts', 'Match Type': 'Broad',
        'full_metrics': fm_ba,
        'Action': {'target_entity': 'Bidding Adjustment',
                   'placement_type': 'Placement Product Page',
                   'set_percentage': 10, 'action': 'Update',
                   'log_msg': ''},
        'IDs': {'Campaign ID': '111', 'Ad Group ID': '222',
                'Keyword ID': '333', 'Ad Group Name': 'AG1',
                'Current_Bid': 0.5, 'EntityType': 'Keyword'},
    },
    {
        'SKU': 'TEST_SKU', 'Campaign Name': 'Camp B',
        'Target': 'gift for nurse', 'Match Type': 'Exact',
        'full_metrics': fm_ba,
        'Action': {'target_entity': 'Keyword', 'action': 'Paused',
                   'log_msg': ''},
        'IDs': {'Campaign ID': '444', 'Ad Group ID': '555',
                'Keyword ID': '666', 'Ad Group Name': 'AG2',
                'Current_Bid': 0.4, 'EntityType': 'Keyword'},
    },
    {
        'SKU': 'TEST_SKU', 'Campaign Name': 'Camp C',
        'Target': 'nurse gifts', 'Match Type': 'Exact',
        'full_metrics': fm_ba,
        'Action': {'target_entity': 'Keyword', 'action': 'Update',
                   'adjust_bid_by': 0.1, 'log_msg': ''},
        'IDs': {'Campaign ID': '777', 'Ad Group ID': '888',
                'Keyword ID': '999', 'Ad Group Name': 'AG3',
                'Current_Bid': 0.56, 'EntityType': 'Keyword'},
    },
]
upload = build_upload_rows(triggered)
assert len(upload) == 3

# Branch A: Bidding Adjustment
r0 = upload[0]
assert r0['Entity']     == 'Bidding Adjustment', r0['Entity']
assert r0['Placement']  == 'Placement Product Page'
assert r0['Percentage'] == '10'
assert r0['Keyword Text'] == ''   # must be blank for BA rows
print("M3 Branch A (Bidding Adjustment): PASS")

# Branch B: Paused keyword
r1 = upload[1]
assert r1['Entity']      == 'Keyword'
assert r1['State']       == 'Paused'
assert r1['Keyword Text']== 'gift for nurse'
assert r1['Placement']   == ''   # must be blank
print("M3 Branch B Paused keyword: PASS")

# Branch B: Bid arithmetic
r2 = upload[2]
assert r2['Bid'] == '0.66', f"Expected 0.66, got {r2['Bid']}"
print("M3 Branch B bid arithmetic (0.56+0.10=0.66): PASS")

print()
print("=== ALL SMOKE TESTS PASSED ===")
