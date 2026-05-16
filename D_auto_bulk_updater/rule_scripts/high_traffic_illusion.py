"""
high_traffic_illusion.py  —  Rule Script
=========================================
Rule   : High_Traffic_Illusion
Trigger: impressions >= 5000  AND  clicks >= 30  AND  orders <= 0
Routing: flat (all match types)
Action : Pause — high impressions + clicks but zero conversion,
         customers browse but drop at checkout
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
    conditions      = rule_config.get('conditions', {})
    impressions_min = conditions.get('impressions_min', 0)
    clicks_min      = conditions.get('clicks_min',      0)
    orders_max      = conditions.get('orders_max',      float('inf'))

    impressions = metrics.get('impressions', 0)
    clicks      = metrics.get('clicks',      0)
    orders      = metrics.get('orders',      0)

    # ── Condition check ───────────────────────────────────────────────────
    if impressions < impressions_min:
        return []
    if clicks < clicks_min:
        return []
    if orders > orders_max:
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
        'action':        rule_config.get('action', 'Paused'),
        'target_entity': rule_config.get('target_entity', 'Keyword'),
        'log_msg':       log_msg,
    }]
