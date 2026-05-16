"""
exact_match_controller.py  —  Rule Script
==========================================
Rule   : Exact_Match_Controller
Trigger: impressions >= 1  AND  clicks >= 1  AND  match_type == Exact
Routing: match_type_logic → 'exact' only

Sub-action priority (evaluated in order):
  1. Winner   : ctr >= winner_ctr_min (20%)  →  no action  (return [])
  2. Pause     : clicks >= pause_click_min (10) AND orders == 0  →  Paused
  3. Bid Cap   : clicks <= low_click_max (4)   AND orders == 0
                 →  set_bid = min(current_bid, low_click_bid_cap)

'sub_action' is interpolated into log_template via .format().
"""


def execute(rule_config: dict, metrics: dict, match_type: str) -> list:
    """
    Parameters
    ----------
    rule_config : full rule object from rules.json
    metrics     : full_metrics dict (11 standard keys + 'current_bid')
    match_type  : e.g. 'Exact', 'Broad', 'Phrase'

    Returns list with one action dict if triggered, else [].
    """
    conditions      = rule_config.get('conditions', {})
    impressions_min = conditions.get('impressions_min', 1)
    clicks_min      = conditions.get('clicks_min',      1)

    impressions = metrics.get('impressions', 0)
    clicks      = metrics.get('clicks',      0)
    orders      = metrics.get('orders',      0)
    ctr         = metrics.get('ctr',         0.0)
    current_bid = metrics.get('current_bid', 0.0)

    # ── Base condition check ──────────────────────────────────────────────
    if impressions < impressions_min:
        return []
    if clicks < clicks_min:
        return []

    # ── Match Type Routing ────────────────────────────────────────────────
    mt_logic = rule_config.get('match_type_logic', {})
    logic    = next(
        (v for k, v in mt_logic.items() if k.lower() == match_type.lower()),
        None
    )
    if logic is None:
        return []

    low_click_max     = logic.get('low_click_max',     4)
    low_click_bid_cap = logic.get('low_click_bid_cap', 0.68)
    pause_click_min   = logic.get('pause_click_min',   10)
    winner_ctr_min    = logic.get('winner_ctr_min',    20.0)

    # ── Sub-action priority evaluation ────────────────────────────────────
    # Priority 1: Winner — high CTR signals strong relevance, no action
    if ctr >= winner_ctr_min:
        return []

    # Priority 2: Fatal spend drain — many clicks, zero orders → Pause
    if clicks >= pause_click_min and orders == 0:
        sub_action = 'Paused'
        payload = {
            'action':        'Paused',
            'target_entity': rule_config.get('target_entity', 'Keyword'),
        }

    # Priority 3: Low engagement — few clicks, zero orders → Bid Cap
    elif clicks <= low_click_max and orders == 0:
        capped = round(min(current_bid, low_click_bid_cap), 2) if current_bid > 0 else low_click_bid_cap
        sub_action = f'Bid Capped (${capped})'
        payload = {
            'action':        'Update',
            'target_entity': rule_config.get('target_entity', 'Keyword'),
            'set_bid':       capped,
        }

    else:
        # Conditions met but no sub-action applies (e.g. clicks in mid-range with orders)
        return []

    # ── Build log message ─────────────────────────────────────────────────
    try:
        log_msg = rule_config.get('log_template', '').format(
            impressions=impressions,
            clicks=clicks,
            orders=orders,
            units=metrics.get('units', 0),
            spend=metrics.get('spend', 0),
            sales=metrics.get('sales', 0),
            acos=metrics.get('acos', 0),
            cpc=metrics.get('cpc', 0),
            ctr=round(ctr, 2),
            cvr=metrics.get('cvr', 0),
            roas=metrics.get('roas', 0),
            match_type=match_type,
            sub_action=sub_action,
        )
    except KeyError:
        log_msg = rule_config.get('log_template', '')

    payload['log_msg'] = log_msg
    return [payload]
