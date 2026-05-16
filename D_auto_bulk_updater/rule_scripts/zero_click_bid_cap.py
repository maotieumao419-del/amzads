"""
zero_click_bid_cap.py  —  Rule Script
======================================
Rule   : Zero_Click_Bid_Cap
Trigger: impressions >= 1  AND  clicks == 0
Routing: flat (all match types)
Action : Cap bid at max_bid_cap (default $0.50).
         Only fires when current_bid > cap (no-op if already within cap).
         Uses set_bid (absolute) so build_upload_rows writes the final value
         directly without adding to current_bid.
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
    clicks_max      = conditions.get('clicks_max',      0)

    impressions  = metrics.get('impressions',  0)
    clicks       = metrics.get('clicks',       0)
    current_bid  = metrics.get('current_bid',  0.0)

    # ── Condition check ───────────────────────────────────────────────────
    if impressions < impressions_min:
        return []
    if clicks > clicks_max:
        return []

    # ── Bid cap logic ─────────────────────────────────────────────────────
    max_cap = float(rule_config.get('max_bid_cap', 0.5))

    # Only act when current_bid is known AND exceeds the cap
    if current_bid <= 0.0 or current_bid <= max_cap:
        return []

    # ── Build payload ─────────────────────────────────────────────────────
    try:
        log_msg = rule_config.get('log_template', '').format(
            impressions=impressions,
            clicks=clicks,
            orders=metrics.get('orders', 0),
            units=metrics.get('units', 0),
            spend=metrics.get('spend', 0),
            sales=metrics.get('sales', 0),
            acos=metrics.get('acos', 0),
            cpc=metrics.get('cpc', 0),
            ctr=metrics.get('ctr', 0),
            cvr=metrics.get('cvr', 0),
            roas=metrics.get('roas', 0),
            match_type=match_type,
            current_bid=current_bid,
            max_bid_cap=max_cap,
        )
    except KeyError:
        log_msg = rule_config.get('log_template', '')

    return [{
        'action':        rule_config.get('action', 'Update'),
        'target_entity': rule_config.get('target_entity', 'Keyword'),
        'set_bid':       max_cap,   # absolute value → build_upload_rows writes directly
        'log_msg':       log_msg,
    }]
