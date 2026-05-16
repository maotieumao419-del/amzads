"""
broad_match_discovery_filter.py  —  Rule Script
================================================
Rule   : Broad_Match_Discovery_Filter
Trigger: impressions >= 1000  AND  clicks <= 2  AND  orders <= 0
Routing: match_type_logic → 'broad' or 'auto'
Action : Pause — discovery waste, customer intent does not match
"""


def execute(rule_config: dict, metrics: dict, match_type: str) -> list:
    """
    Parameters
    ----------
    rule_config : full rule object from rules.json
    metrics     : full_metrics dict (11 keys)
    match_type  : e.g. 'Exact', 'Broad', 'Phrase', 'auto'

    Returns list with one action dict if triggered, else [].
    """
    conditions      = rule_config.get('conditions', {})
    impressions_min = conditions.get('impressions_min', 0)
    clicks_max      = conditions.get('clicks_max',      float('inf'))
    orders_max      = conditions.get('orders_max',      float('inf'))

    impressions = metrics.get('impressions', 0)
    clicks      = metrics.get('clicks',      0)
    orders      = metrics.get('orders',      0)

    # ── Condition check ───────────────────────────────────────────────────
    if impressions < impressions_min:
        return []
    if clicks > clicks_max:
        return []
    if orders > orders_max:
        return []

    # ── Match Type Routing ────────────────────────────────────────────────
    mt_logic = rule_config.get('match_type_logic', {})
    logic    = next(
        (v for k, v in mt_logic.items() if k.lower() == match_type.lower()),
        None
    )
    if logic is None:
        return []

    # ── Build payload ─────────────────────────────────────────────────────
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
            ctr=metrics.get('ctr', 0),
            cvr=metrics.get('cvr', 0),
            roas=metrics.get('roas', 0),
            match_type=match_type,
        )
    except KeyError:
        log_msg = rule_config.get('log_template', '')

    return [{
        'action':        logic.get('action', 'Paused'),
        'target_entity': logic.get('target_entity', 'Keyword'),
        'log_msg':       log_msg,
    }]
