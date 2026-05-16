"""
auto_bulk_updater.py  —  Core Logic Layer  (Strategy-Pattern Edition)
======================================================================
Role  : Reads PPC_*_UPDATED.json + BulkSheetExport_*.json, runs the
        Dynamic Rule Engine, and generates Amazon Upload bulksheets.

Input  : data/working_json/PPC_*_UPDATED.json
         data/working_json/BulkSheetExport_*.json
         rules.json
         rule_scripts/<rule_name_lowercase>.py  (optional dynamic modules)

Output : data/working_json/Amazon_Upload_Ready_*.json
         data/final_xlsx/Amazon_Upload_Ready_*.xlsx

Architecture: Strategy Pattern
  Each rule in rules.json CAN have a corresponding module in rule_scripts/.
  If the module exists  →  call module.execute(rule_config, metrics, match_type)
  If the module is absent →  static fallback evaluator runs instead.
  This means the engine NEVER crashes on a missing rule script.

── Module 1: 11 Full-Funnel Metrics ─────────────────────────────────────────
  JSON Key               full_metrics key   Type   Fallback formula
  ─────────────────────  ────────────────   ─────  ──────────────────────────
  "Impressions"        → impressions        int
  "Clicks"             → clicks             int
  "Click-through Rate" → ctr               float   clicks / impressions × 100
  "Spend"              → spend             float
  "Sales"              → sales             float
  "Orders"             → orders             int
  "Units"              → units              int
  "Conversion Rate"    → cvr               float   orders / clicks × 100
  "ACOS"               → acos              float   spend / sales × 100
  "CPC"                → cpc               float   spend / clicks
  "ROAS"               → roas              float   sales / spend
─────────────────────────────────────────────────────────────────────────────
"""

import os
import sys
import json
import glob
import logging
import re
import importlib.util
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, 'io_handlers'))
from excel_json_handler import json_to_xlsx, AMAZON_BULK_COLUMNS

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

JSON_DIR      = os.path.join(BASE_DIR, 'data', 'working_json')
FINAL_DIR     = os.path.join(BASE_DIR, 'data', 'final_xlsx')
RULES_FILE    = os.path.join(BASE_DIR, 'rules.json')
SCRIPTS_DIR   = os.path.join(BASE_DIR, 'rule_scripts')
os.makedirs(FINAL_DIR, exist_ok=True)

_MATCH_RE   = re.compile(r'\[\[(Exact|Phrase|Broad)\]', re.IGNORECASE)
_GARBAGE_RE = re.compile(r'[^\d.]')   # strips $, %, comma, spaces …

# Cache for dynamically loaded rule modules  {rule_name_lower: module | None}
_MODULE_CACHE: dict = {}


# ═══════════════════════════════════════════════════════════════════════════
# MODULE 1 — Data Pipeline & Full-Funnel Metrics Extraction
# ═══════════════════════════════════════════════════════════════════════════

def clean_metric(value, type_func=float):
    """
    Strip garbage characters ($, %, comma …) from a raw JSON value
    and safely cast to type_func (float or int).
    Returns 0 of the requested type on any error or empty input.

    Examples
    --------
    clean_metric("$1,234.56")        → 1234.56
    clean_metric("12.50%", float)    → 12.5
    clean_metric("1,200",  int)      → 1200
    clean_metric(None)               → 0.0
    clean_metric("--")               → 0.0
    """
    try:
        if value is None or str(value).strip() in ('', '-', '--', 'N/A', 'n/a'):
            return type_func(0)
        cleaned = _GARBAGE_RE.sub('', str(value))
        if not cleaned:
            return type_func(0)
        return type_func(float(cleaned))
    except (ValueError, TypeError):
        return type_func(0)


def extract_full_metrics(raw_metrics: dict) -> dict:
    """
    Parse + interpolate all 11 full-funnel metrics from one metrics_by_date
    object. Rate fields that are 0 but derivable from base values are
    auto-computed (fallback mechanism).
    """
    # ── Absolute base values ────────────────────────────────────────────────
    impressions = clean_metric(raw_metrics.get('Impressions', 0),       int)
    clicks      = clean_metric(raw_metrics.get('Clicks', 0),            int)
    spend       = clean_metric(raw_metrics.get('Spend', 0),             float)
    sales       = clean_metric(raw_metrics.get('Sales', 0),             float)
    orders      = clean_metric(raw_metrics.get('Orders', 0),            int)
    units       = clean_metric(raw_metrics.get('Units', 0),             int)

    # ── Rate / Ratio values with Fallback Interpolation ─────────────────────
    # CTR = Clicks / Impressions × 100
    ctr = clean_metric(raw_metrics.get('Click-through Rate', 0), float)
    if ctr == 0.0 and impressions > 0 and clicks > 0:
        ctr = round((clicks / impressions) * 100, 4)

    # CVR = Orders / Clicks × 100
    cvr = clean_metric(raw_metrics.get('Conversion Rate', 0), float)
    if cvr == 0.0 and clicks > 0 and orders > 0:
        cvr = round((orders / clicks) * 100, 4)

    # ACOS = Spend / Sales × 100
    acos = clean_metric(raw_metrics.get('ACOS', 0), float)
    if acos == 0.0 and sales > 0 and spend > 0:
        acos = round((spend / sales) * 100, 4)

    # CPC = Spend / Clicks
    cpc = clean_metric(raw_metrics.get('CPC', 0), float)
    if cpc == 0.0 and clicks > 0 and spend > 0:
        cpc = round(spend / clicks, 4)

    # ROAS = Sales / Spend
    roas = clean_metric(raw_metrics.get('ROAS', 0), float)
    if roas == 0.0 and spend > 0 and sales > 0:
        roas = round(sales / spend, 4)

    return {
        'impressions': impressions,
        'clicks':      clicks,
        'ctr':         ctr,
        'spend':       spend,
        'sales':       sales,
        'orders':      orders,
        'units':       units,
        'cvr':         cvr,
        'acos':        acos,
        'cpc':         cpc,
        'roas':        roas,
    }


# ═══════════════════════════════════════════════════════════════════════════
# MODULE 2 — Dynamic Rule Engine (Strategy Pattern)
# ═══════════════════════════════════════════════════════════════════════════

def load_rules() -> dict:
    """Load rules.json; write a minimal default if the file is missing."""
    if os.path.exists(RULES_FILE):
        try:
            with open(RULES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"Error reading rules.json: {e}")

    default = {
        "rules": [
            {
                "rule_name": "Fatal_Bleeder_Zero_Sales",
                "conditions": {"clicks_min": 15, "orders_max": 0},
                "action": "Paused",
                "target_entity": "Keyword",
                "log_template": "Paused (Fatal Bleeder): {clicks} Clicks, {orders} Orders."
            }
        ]
    }
    logging.warning("rules.json not found — writing default fallback.")
    with open(RULES_FILE, 'w', encoding='utf-8') as f:
        json.dump(default, f, indent=2)
    return default


def _load_rule_module(rule_name: str):
    """
    Attempt to import rule_scripts/<rule_name_lower>.py via importlib.
    Returns the module object on success, or None if the file doesn't exist.
    Results are cached in _MODULE_CACHE to avoid repeated filesystem hits.
    """
    key = rule_name.lower()
    if key in _MODULE_CACHE:
        return _MODULE_CACHE[key]

    module_path = os.path.join(SCRIPTS_DIR, f"{key}.py")
    if not os.path.isfile(module_path):
        _MODULE_CACHE[key] = None
        return None

    try:
        spec   = importlib.util.spec_from_file_location(key, module_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        _MODULE_CACHE[key] = module
        logging.info(f"[RuleEngine] Loaded dynamic module: {key}.py")
        return module
    except Exception as e:
        logging.warning(f"[RuleEngine] Failed to load {key}.py: {e}")
        _MODULE_CACHE[key] = None
        return None


def _static_evaluate(rule: dict, full_metrics: dict, match_type: str) -> list:
    """
    Built-in condition evaluator used when no dynamic rule script exists.
    Supports all 11 metric condition keys (_min / _max).
    Handles both flat action dicts and match_type_logic branching.
    """
    cond  = rule.get('conditions', {})
    m     = full_metrics

    checks = [
        ('impressions_min', m['impressions'], False),
        ('impressions_max', m['impressions'], True),
        ('clicks_min',      m['clicks'],      False),
        ('clicks_max',      m['clicks'],      True),
        ('orders_min',      m['orders'],      False),
        ('orders_max',      m['orders'],      True),
        ('units_min',       m['units'],       False),
        ('units_max',       m['units'],       True),
        ('spend_min',       m['spend'],       False),
        ('spend_max',       m['spend'],       True),
        ('sales_min',       m['sales'],       False),
        ('sales_max',       m['sales'],       True),
        ('acos_min',        m['acos'],        False),
        ('acos_max',        m['acos'],        True),
        ('cpc_min',         m['cpc'],         False),
        ('cpc_max',         m['cpc'],         True),
        ('ctr_min',         m['ctr'],         False),
        ('ctr_max',         m['ctr'],         True),
        ('cvr_min',         m['cvr'],         False),
        ('cvr_max',         m['cvr'],         True),
        ('roas_min',        m['roas'],        False),
        ('roas_max',        m['roas'],        True),
    ]
    for key, val, is_max in checks:
        if key not in cond:
            continue
        if is_max and val > cond[key]:
            return []
        if not is_max and val < cond[key]:
            return []

    # Conditions passed — build log message
    try:
        log_msg = rule.get('log_template', '').format(
            impressions=m['impressions'], clicks=m['clicks'],
            orders=m['orders'],          units=m['units'],
            spend=m['spend'],            sales=m['sales'],
            acos=m['acos'],              cpc=m['cpc'],
            ctr=m['ctr'],               cvr=m['cvr'],
            roas=m['roas'],             match_type=match_type
        )
    except KeyError:
        log_msg = rule.get('log_template', '')

    # ── match_type_logic branch (per-match-type overrides) ─────────────────
    mt_logic = rule.get('match_type_logic', {})
    if mt_logic:
        logic = next(
            (v for k, v in mt_logic.items() if k.lower() == match_type.lower()),
            None
        )
        if not logic:
            return []          # rule has MT logic but this match_type isn't listed
        result = logic.copy()
        result['log_msg'] = log_msg
        return [result]

    # ── Flat rule (applies to all match types) ──────────────────────────────
    result = {
        'action':        rule.get('action', 'Update'),
        'target_entity': rule.get('target_entity', 'Keyword'),
        'log_msg':       log_msg,
    }
    # Carry optional Bidding Adjustment fields if present at rule level
    for field in ('placement_type', 'set_percentage', 'adjust_bid_by'):
        if field in rule:
            result[field] = rule[field]
    return [result]


def run_rule_engine(full_metrics: dict, match_type: str,
                    rules_config: dict) -> list:
    """
    Iterate over every rule in rules_config.

    For each rule:
      1. Try to load rule_scripts/<rule_name_lower>.py via importlib.
      2. If the module exists and has execute(), call:
             module.execute(rule_config, metrics=full_metrics, match_type=match_type)
         The module must return a list of action dicts (empty list = no action).
      3. If the module is missing (ModuleNotFoundError) or lacks execute()
         (AttributeError), fall back silently to _static_evaluate().

    Returns list of triggered action dicts across all rules.
    """
    triggered_actions = []

    for rule in rules_config.get('rules', []):
        rule_name = rule.get('rule_name', '')
        module    = _load_rule_module(rule_name)

        if module is not None:
            # ── Dynamic path ───────────────────────────────────────────────
            try:
                results = module.execute(
                    rule_config=rule,
                    metrics=full_metrics,
                    match_type=match_type
                )
                if isinstance(results, list):
                    triggered_actions.extend(results)
            except ModuleNotFoundError as e:
                logging.warning(
                    f"[RuleEngine] ModuleNotFoundError in '{rule_name}': {e} "
                    f"— falling back to static evaluator."
                )
                triggered_actions.extend(
                    _static_evaluate(rule, full_metrics, match_type)
                )
            except AttributeError as e:
                logging.warning(
                    f"[RuleEngine] AttributeError in '{rule_name}' "
                    f"(missing execute()?): {e} — falling back to static evaluator."
                )
                triggered_actions.extend(
                    _static_evaluate(rule, full_metrics, match_type)
                )
            except Exception as e:
                logging.error(
                    f"[RuleEngine] Unexpected error in '{rule_name}': {e}"
                )
        else:
            # ── Static fallback path ───────────────────────────────────────
            triggered_actions.extend(
                _static_evaluate(rule, full_metrics, match_type)
            )

    return triggered_actions


# ═══════════════════════════════════════════════════════════════════════════
# ID Lookup Map from Bulk JSON
# ═══════════════════════════════════════════════════════════════════════════

def build_id_map(bulk_records: list) -> tuple:
    """Returns (campaign_map, keyword_map)."""
    campaign_map = {}
    keyword_map  = {}

    for row in bulk_records:
        entity  = str(row.get('Entity', '')).strip().lower()
        n_camp  = str(row.get('Campaign Name', '')).strip().lower()
        n_mt    = str(row.get('Match Type', '')).strip().lower()
        if entity == 'product targeting' and not n_mt:
            n_mt = 'targeting'

        raw_kw  = row.get('Keyword Text') or row.get('Product Targeting Expression') or ''
        n_kw    = str(raw_kw).strip().lower()

        c_id    = str(row.get('Campaign ID', '')).strip().rstrip('.0')
        ag_id   = str(row.get('Ad Group ID', '')).strip().rstrip('.0')
        kw_id   = str(row.get('Keyword ID', '')).strip().rstrip('.0')
        ag_name = str(row.get('Ad Group Name', '')).strip()

        try:
            current_bid = float(row.get('Bid', 0) or 0)
        except (ValueError, TypeError):
            current_bid = 0.0

        if n_camp and c_id:
            campaign_map[n_camp] = c_id

        if entity in ('keyword', 'product targeting') and n_camp and (n_kw or n_mt):
            keyword_map[(n_camp, n_kw, n_mt)] = {
                'Campaign ID':   c_id,
                'Ad Group ID':   ag_id,
                'Keyword ID':    kw_id,
                'Ad Group Name': ag_name,
                'Current_Bid':   current_bid,
                'EntityType':    'Product Targeting' if entity == 'product targeting' else 'Keyword',
            }

    return campaign_map, keyword_map


# ═══════════════════════════════════════════════════════════════════════════
# Process Internal JSON
# ═══════════════════════════════════════════════════════════════════════════

def process_internal(internal_data: dict, campaign_map: dict,
                     keyword_map: dict, rules_config: dict) -> list:
    """
    Walk every SKU sheet → every row → extract full_metrics → run_rule_engine.
    Returns a flat list of triggered-action dicts.
    """
    triggered = []

    for sheet_name, sheet_data in internal_data.items():
        if sheet_name in ('Listing', 'Portfolio ID') or not isinstance(sheet_data, dict):
            continue
        if not sheet_data.get('date_blocks'):
            continue

        last_date = sheet_data['date_blocks'][-1]

        for row in sheet_data.get('rows', []):
            camp_name = str(row.get('Campaign Name', '')).strip()
            target    = str(row.get('Target', '')).strip()
            if not camp_name or not target:
                continue

            note = str(row.get('Ghi chú', ''))
            mt_match   = _MATCH_RE.search(note)
            match_type = mt_match.group(1) if mt_match else ''

            # Resolve IDs early — needed to inject current_bid into full_metrics
            n_camp = camp_name.lower()
            n_kw   = target.lower()
            n_mt   = match_type.lower().strip('[]')
            ids    = keyword_map.get((n_camp, n_kw, n_mt), {})

            # Module 1 — extract & interpolate 11 metrics + current_bid
            raw_metrics  = row.get('metrics_by_date', {}).get(last_date, {})
            full_metrics = extract_full_metrics(raw_metrics)
            try:
                full_metrics['current_bid'] = float(ids.get('Current_Bid', 0.0))
            except (ValueError, TypeError):
                full_metrics['current_bid'] = 0.0

            # Module 2 — dynamic rule engine
            actions = run_rule_engine(full_metrics, match_type, rules_config)

            for act in actions:
                if not ids:
                    logging.warning(
                        f"ID Lookup Failed: Camp='{n_camp}', "
                        f"Target='{n_kw}', Match='{n_mt}'"
                    )

                triggered.append({
                    'SKU':           sheet_name,
                    'Campaign Name': camp_name,
                    'Target':        target,
                    'Match Type':    match_type,
                    'full_metrics':  full_metrics,
                    'Action':        act,
                    'IDs':           ids,
                })

    return triggered


# ═══════════════════════════════════════════════════════════════════════════
# MODULE 3 — Bulk File Output Formatter
# ═══════════════════════════════════════════════════════════════════════════

def build_upload_rows(triggered_actions: list) -> list:
    """
    Convert triggered-action dicts into Amazon Advertising Bulksheet rows.

    Entity routing
    ──────────────
    Priority for entity type:
      1. act['target_entity']         (returned by rule engine)
      2. ids['EntityType']            (from ID map)
      3. fallback → 'Keyword'

    Branch A — "Bidding Adjustment"
      Only fills: Operation, Placement, Percentage.
      All keyword/targeting columns are left blank.

    Branch B — "Keyword" / "Product Targeting"
      Fills: Campaign/AdGroup/Keyword IDs, State, Keyword or Targeting text.
      Bid adjustment: Current_Bid + adjust_bid_by  (math, not string concat).
    """
    rows = []

    for item in triggered_actions:
        act  = item['Action']
        ids  = item.get('IDs', {})

        # ── Resolve entity type (priority order) ───────────────────────────
        ent_type = (
            act.get('target_entity')
            or ids.get('EntityType')
            or 'Keyword'
        )

        row = {col: '' for col in AMAZON_BULK_COLUMNS}
        row['Product']       = 'Sponsored Products'
        row['Entity']        = ent_type
        row['Operation']     = 'Update'
        row['Campaign Name'] = item['Campaign Name']
        row['Campaign ID']   = ids.get('Campaign ID', '')

        # ══════════════════════════════════════════════════════════════════
        # Branch A — Bidding Adjustment
        # ══════════════════════════════════════════════════════════════════
        if ent_type.lower() == 'bidding adjustment':
            row['Placement']  = act.get('placement_type', '')
            try:
                row['Percentage'] = str(int(float(act.get('set_percentage', 0))))
            except (ValueError, TypeError):
                row['Percentage'] = '0'
            rows.append(row)
            continue

        # ══════════════════════════════════════════════════════════════════
        # Branch B — Keyword / Product Targeting
        # ══════════════════════════════════════════════════════════════════
        row['Ad Group ID']   = ids.get('Ad Group ID', '')
        row['Keyword ID']    = ids.get('Keyword ID', '')
        row['Ad Group Name'] = ids.get('Ad Group Name', '')

        # State (Paused / Enabled / Archived)
        action_val = str(act.get('action', 'Update')).strip()
        if action_val.lower() in ('paused', 'enabled', 'archived'):
            row['State'] = action_val.title()

        # Keyword text vs Product Targeting Expression
        if ent_type.lower() == 'product targeting':
            row['Product Targeting Expression'] = item['Target']
            row['Match Type'] = ''
        else:
            row['Keyword Text'] = item['Target']
            row['Match Type']   = item['Match Type']

        # Bid — priority: set_bid (absolute cap) > adjust_bid_by (relative delta)
        if 'set_bid' in act:
            # Absolute bid value supplied directly by a rule script (e.g. bid-cap rules)
            try:
                row['Bid'] = str(round(float(act['set_bid']), 2))
            except (ValueError, TypeError):
                pass
        elif 'adjust_bid_by' in act:
            base_bid = ids.get('Current_Bid', 0.0)
            try:
                adj = float(act['adjust_bid_by'])
            except (ValueError, TypeError):
                adj = 0.0
            new_bid = round(base_bid + adj, 2)
            row['Bid'] = str(new_bid)

        rows.append(row)

    return rows


# ═══════════════════════════════════════════════════════════════════════════
# Entry Point
# ═══════════════════════════════════════════════════════════════════════════

def main():
    print("--- AUTO BULK UPDATER  [Strategy-Pattern Edition] ---")
    rules_config = load_rules()
    today        = datetime.now().strftime('%d%m%Y')

    # 1. Build ID map from Bulk JSON files
    campaign_map, keyword_map = {}, {}
    for bj in sorted(glob.glob(os.path.join(JSON_DIR, 'BulkSheetExport_*.json'))):
        with open(bj, 'r', encoding='utf-8') as f:
            records = json.load(f)
        cm, km = build_id_map(records)
        campaign_map.update(cm)
        keyword_map.update(km)
    logging.info(f"ID Map: {len(campaign_map)} campaigns, {len(keyword_map)} keywords.")

    # 2. Process every UPDATED Internal JSON
    updated_jsons = sorted(glob.glob(os.path.join(JSON_DIR, 'PPC_*_UPDATED.json')))
    if not updated_jsons:
        logging.error("No PPC_*_UPDATED.json found. Run end_to_end_ppc_tracker.py first.")
        return

    all_triggered = []
    for uj in updated_jsons:
        with open(uj, 'r', encoding='utf-8') as f:
            internal_data = json.load(f)
        triggered = process_internal(
            internal_data, campaign_map, keyword_map, rules_config
        )
        all_triggered.extend(triggered)
        logging.info(f"  {os.path.basename(uj)}: {len(triggered)} actions triggered.")

    # 3. Generate Amazon Upload JSON + xlsx
    if all_triggered:
        upload_rows   = build_upload_rows(all_triggered)
        out_json_name = f"Amazon_Upload_Ready_{today}.json"
        out_json_path = os.path.join(JSON_DIR, out_json_name)
        with open(out_json_path, 'w', encoding='utf-8') as f:
            json.dump(upload_rows, f, ensure_ascii=False, indent=2)
        logging.info(f"Upload JSON: {out_json_name} ({len(upload_rows)} rows)")

        out_xlsx = os.path.join(FINAL_DIR, out_json_name.replace('.json', '.xlsx'))
        if json_to_xlsx(out_json_path, out_xlsx):
            logging.info(f"Upload xlsx: {os.path.basename(out_xlsx)}")
    else:
        logging.info("No rules triggered. No upload file generated.")

    logging.info("UPDATER COMPLETE.")


if __name__ == "__main__":
    main()
