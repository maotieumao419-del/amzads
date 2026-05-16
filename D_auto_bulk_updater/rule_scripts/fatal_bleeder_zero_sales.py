"""
fatal_bleeder_zero_sales.py  —  Rule Script
============================================
Rule   : Fatal_Bleeder_Zero_Sales
Trigger: clicks >= 15  AND  orders <= 0
Routing: flat (all match types)
Action : Pause keyword immediately — high spend, zero conversion
"""


def execute(rule_config: dict, metrics: dict, match_type: str) -> list:
    """
    Parameters
    ----------
    rule_config : full rule object from rules.json
    metrics     : full_metrics dict (11 keys)
    match_type  : e.g. 'Exact', 'Broad', 'Phrase'

    Returns list with one action dict if triggered, else [].
    """
    conditions = rule_config.get('conditions', {})
    clicks_min  = conditions.get('clicks_min',  0)
    orders_max  = conditions.get('orders_max',  float('inf'))

    clicks = metrics.get('clicks', 0)
    orders = metrics.get('orders', 0)

    # ── Condition check ───────────────────────────────────────────────────
    if clicks < clicks_min:
        return []
    if orders > orders_max:
        return []

    # ── Build payload (flat rule — no match_type_logic) ───────────────────
    try:
        log_msg = rule_config.get('log_template', '').format(
            impressions=metrics.get('impressions', 0),
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
        'action':        rule_config.get('action', 'Paused'),
        'target_entity': rule_config.get('target_entity', 'Keyword'),
        'log_msg':       log_msg,
    }]
